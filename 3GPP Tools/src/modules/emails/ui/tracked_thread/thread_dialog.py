"""Main TrackedThreadDialog composition and lifecycle handlers."""

from pathlib import Path
from typing import Callable, Dict, List, Optional

from PyQt5.QtCore import QSettings, Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import QDialog, QWidget

from modules.emails.core.ai_summary import EmailSummaryStore
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_workers import SingleThreadRefreshWorker

from .mixins import (
    TrackedThreadActionsMixin,
    TrackedThreadLinksNotesMixin,
    TrackedThreadRenderingMixin,
    TrackedThreadUiSetupMixin,
)


class TrackedThreadDialog(
    TrackedThreadUiSetupMixin,
    TrackedThreadRenderingMixin,
    TrackedThreadLinksNotesMixin,
    TrackedThreadActionsMixin,
    QDialog,
):
    """Integrated viewer for a tracked email thread with Markdown export and LLM CLI launch."""

    thread_changed = pyqtSignal()

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        thread_id: str,
        parent: Optional[QWidget] = None,
        ai_suggestions: Optional[List[str]] = None,
        wi_suggestions: Optional[List[str]] = None,
        tdoc_suggestions: Optional[List[str]] = None,
        tdoc_family_expander: Optional[Callable[[List[str]], List[str]]] = None,
        sub_item_suggestions: Optional[List[str]] = None,
        all_existing_tags: Optional[List[str]] = None,
        annotations_save_callback: Optional[Callable[[str, str, str, str], None]] = None,
        ai_to_sub_items: Optional[Dict[str, List[str]]] = None,
        open_tdoc_callback: Optional[Callable[[str], None]] = None,
    ):
        super().__init__(
            parent,
            Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint,
        )
        self.db = db
        self.thread_id = thread_id
        self._ai_suggestions = ai_suggestions or []
        self._wi_suggestions = wi_suggestions or []
        self._tdoc_suggestions = tdoc_suggestions or []
        self._sub_item_suggestions = sub_item_suggestions or []
        self._ai_to_sub_items: Dict[str, List[str]] = self._normalize_ai_sub_map(
            ai_to_sub_items, self._sub_item_suggestions
        )
        self._all_existing_tags = all_existing_tags or []
        self._annotations_save_callback = annotations_save_callback
        self._tdoc_family_expander = tdoc_family_expander
        self._open_tdoc_callback = open_tdoc_callback
        self._loading = False
        self._refresh_worker: Optional[SingleThreadRefreshWorker] = None
        self._unread_blocks: List[object] = []
        self._unread_cursor: int = -1
        self._initial_scroll_done: bool = False
        db_p = getattr(db, "db_path", None)
        self._summary_store: Optional[EmailSummaryStore] = (
            EmailSummaryStore(Path(db_p).parent / "email_summaries.db") if db_p else None
        )
        self._settings = QSettings("3GPP_Tools", "TrackedThreadDialog")
        self.setWindowFlags(Qt.Window)
        self.setSizeGripEnabled(True)
        self.setWindowTitle("Tracked Email Thread")
        self.setMinimumSize(520, 420)
        self.resize(1020, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass
        self._build_ui()
        self._load_data()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._initial_scroll_done:
            QTimer.singleShot(50, self._scroll_to_initial_unread)


    def closeEvent(self, event):
        try:
            if hasattr(self, "summary_card"):
                self.summary_card.shutdown_worker()
            if getattr(self, "_refresh_worker", None) and self._refresh_worker.isRunning():
                self._refresh_worker.request_cancel()
                self._refresh_worker.wait(1500)
        except Exception:
            pass
        try:
            self._settings.setValue("geometry", self.saveGeometry())
            if hasattr(self, "browser"):
                self._settings.setValue("zoom_steps", int(self.browser.property("_reader_zoom_steps") or 0))
        except Exception:
            pass
        super().closeEvent(event)

