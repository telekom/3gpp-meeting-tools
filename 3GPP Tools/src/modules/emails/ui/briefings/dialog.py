"""Modeless EmailBriefingsDialog window for Per-AI cross-thinking & Saved Selections."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence

from PyQt5.QtCore import QSettings, Qt, pyqtSignal
from PyQt5.QtWidgets import QDialog, QTreeWidgetItem

from core.ai.llm_runner import ModelSpec
from modules.emails.core.ai_summary import EmailSummaryRecord, EmailSummaryStore

from .mixins import BriefingsActionsMixin, BriefingsUiSetupMixin


class EmailBriefingsDialog(BriefingsUiSetupMixin, BriefingsActionsMixin, QDialog):
    """Independent modeless top-level window for Per-AI email briefings and saved selections."""

    briefing_saved = pyqtSignal(object)  # EmailSummaryRecord

    def __init__(
        self,
        *,
        tracked_db: Optional[Any] = None,
        general_email_db: Optional[Any] = None,
        summary_db_path: Optional[Path] = None,
        meeting_id: str = "",
        wg: str = "SA2",
        ai_label_map: Optional[Dict[str, str]] = None,
        followed_ais: Optional[Iterable[str]] = None,
        tdoc_ai_lookup: Optional[Dict[str, str]] = None,
        tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
        open_tdoc_callback: Optional[Callable[[str], None]] = None,
        open_thread_callback: Optional[Callable[[str], None]] = None,
        jump_tdoc_callback: Optional[Callable[[str], None]] = None,
        initial_ai_code: str = "",
        initial_selection: Optional[Dict[str, Any]] = None,
        auto_run_initial: bool = False,
        parent=None,
    ) -> None:
        super().__init__(None)  # None avoids Win32 GW_OWNER pinning
        if parent is not None:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass

        self.tracked_db = tracked_db
        self.general_email_db = general_email_db
        self.meeting_id = str(meeting_id or "").strip()
        self.wg = str(wg or "SA2").strip().upper()
        self.ai_label_map = dict(ai_label_map or {})
        self.followed_ais = {str(a).strip() for a in (followed_ais or ()) if str(a).strip()}
        self.tdoc_ai_lookup = dict(tdoc_ai_lookup or {})
        self.tdoc_meta_map = dict(tdoc_meta_map or {})
        self._open_tdoc_callback = open_tdoc_callback
        self._open_thread_callback = open_thread_callback
        self._jump_tdoc_callback = jump_tdoc_callback

        resolved_db = summary_db_path
        if resolved_db is None:
            base_p = getattr(tracked_db, "db_path", None) or getattr(general_email_db, "db_path", None)
            if base_p:
                resolved_db = Path(base_p).parent / "email_summaries.db"
            else:
                resolved_db = Path.home() / "3GPP_Delegate_Helper" / "temp" / "email_summaries.db"
        self.summary_store = EmailSummaryStore(Path(resolved_db))
        self.meeting_dir = self.summary_store.db_path.parent.parent

        self._ai_buckets: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self._active_record: Optional[EmailSummaryRecord] = None
        self._synth_model_override: Optional[ModelSpec] = None
        self._worker = None
        self._m365_dialog = None
        self._stream_chunks: List[str] = []

        self._pending_selection_key = ""
        self._pending_selection_title = ""
        self._pending_selection_messages: List[Dict[str, Any]] = []

        self._settings = QSettings("3GPP_Tools", "EmailBriefingsDialog")
        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        suffix = f" — {self.meeting_id}" if self.meeting_id else ""
        self.setWindowTitle(f"🤖 AI Email Briefings{suffix}")
        self.setMinimumSize(720, 480)
        self.resize(1120, 740)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        self._build_ui()

        if initial_selection and initial_selection.get("messages"):
            self.start_selection_briefing(
                title=str(initial_selection.get("title") or "Custom Selection"),
                messages=initial_selection.get("messages") or [],
                focus_prompt=str(initial_selection.get("focus_prompt") or ""),
                auto_run=auto_run_initial,
            )
        else:
            self.refresh_navigation(
                select_kind="ai" if initial_ai_code else "",
                select_key=str(initial_ai_code or "").strip(),
            )
            if initial_ai_code and auto_run_initial:
                self._run_active_briefing(mode="auto")

    def focus_ai(self, ai_code: str, *, auto_run: bool = False) -> None:
        self.refresh_navigation(select_kind="ai", select_key=str(ai_code or "").strip())
        if auto_run:
            self._run_active_briefing(mode="auto")

    def start_selection_briefing(
        self,
        *,
        title: str,
        messages: Sequence[Dict[str, Any]],
        focus_prompt: str = "",
        auto_run: bool = True,
    ) -> None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sel_key = f"SEL-{stamp}"
        self._pending_selection_key = sel_key
        self._pending_selection_title = title.strip() or sel_key
        self._pending_selection_messages = list(messages or [])

        self.refresh_navigation(select_kind="selection", select_key=sel_key)
        # If not yet persisted, inject temporary selection item & select it
        if not self.summary_store.get_summary("selection", sel_key):
            root = self.nav_tree.invisibleRootItem()
            sel_root = root.child(1) if root.childCount() > 1 else None
            if sel_root is not None:
                item = QTreeWidgetItem(sel_root, [f"✨ {self._pending_selection_title} ({len(self._pending_selection_messages)})"])
                item.setData(0, Qt.UserRole, ("selection", sel_key))
                self.nav_tree.setCurrentItem(item)
                self._load_selected_scope()
        if focus_prompt:
            self.focus_prompt_edit.setText(focus_prompt)
        if auto_run and self._pending_selection_messages:
            self._run_active_briefing(mode="full")

    def closeEvent(self, event) -> None:
        try:
            if self._worker and self._worker.isRunning():
                self._worker.request_cancel()
                self._worker.wait(1000)
            if self._m365_dialog is not None:
                self._m365_dialog.close()
                self._m365_dialog = None
            self._settings.setValue("geometry", self.saveGeometry())
        except Exception:
            pass
        super().closeEvent(event)
