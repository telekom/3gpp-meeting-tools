"""Background QThread workers wrapping `EmailSummaryOrchestrator` with cooperative cancellation."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PyQt5.QtCore import QThread, pyqtSignal

from core.ai.llm_runner import CancellationFlag, ModelSpec

from .models import SummaryScope
from .orchestrator import EmailSummaryOrchestrator
from .store import EmailSummaryStore


class EmailSummaryTaskWorker(QThread):
    """QThread executing a single Thread, TDoc, or Selection summary task."""

    progress_changed = pyqtSignal(str)
    token_received = pyqtSignal(str)
    summary_completed = pyqtSignal(object)  # EmailSummaryRecord
    summary_failed = pyqtSignal(str)

    def __init__(
        self,
        db_path: Path,
        scope: SummaryScope,
        messages: Sequence[Dict[str, Any]],
        *,
        meeting_dir: Optional[Path] = None,
        wg: str = "",
        meeting_id: str = "",
        mode: str = "auto",
        model_spec: Optional[ModelSpec] = None,
        tdoc_meta: Optional[Dict[str, Any]] = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.scope = scope
        self.messages = list(messages or [])
        self.meeting_dir = Path(meeting_dir) if meeting_dir else self.db_path.parent.parent
        self.wg = wg
        self.meeting_id = meeting_id
        self.mode = mode
        self.model_spec = model_spec
        self.tdoc_meta = tdoc_meta
        self._cancel_flag = CancellationFlag()

    def request_cancel(self) -> None:
        self._cancel_flag.cancel()
        self.requestInterruption()

    def run(self) -> None:
        try:
            store = EmailSummaryStore(self.db_path)
            orch = EmailSummaryOrchestrator(
                store,
                self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
            )
            record, err = orch.summarize_leaf(
                self.scope,
                self.messages,
                mode=self.mode,
                model_spec=self.model_spec,
                tdoc_meta=self.tdoc_meta,
                cancel_flag=self._cancel_flag,
                on_progress=self.progress_changed.emit,
                on_chunk=self.token_received.emit,
            )
            if self._cancel_flag.is_cancelled() or self.isInterruptionRequested():
                self.summary_failed.emit("Summary cancelled by user.")
                return
            if err or record is None:
                self.summary_failed.emit(err or "Summary generation failed.")
                return
            self.summary_completed.emit(record)
        except Exception as exc:
            if not self._cancel_flag.is_cancelled():
                self.summary_failed.emit(str(exc))


class AiBriefingWorker(QThread):
    """QThread executing 2-stage Map-Reduce Per-AI cross-thinking synthesis."""

    progress_changed = pyqtSignal(str)
    token_received = pyqtSignal(str)
    summary_completed = pyqtSignal(object)  # EmailSummaryRecord
    summary_failed = pyqtSignal(str)

    def __init__(
        self,
        db_path: Path,
        ai_code: str,
        ai_label: str,
        leaf_buckets: Dict[str, Dict[str, Any]],
        *,
        meeting_dir: Optional[Path] = None,
        wg: str = "",
        meeting_id: str = "",
        mode: str = "auto",
        leaf_model_spec: Optional[ModelSpec] = None,
        synthesis_model_spec: Optional[ModelSpec] = None,
        tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
        focus_prompt: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.ai_code = ai_code
        self.ai_label = ai_label
        self.leaf_buckets = dict(leaf_buckets or {})
        self.meeting_dir = Path(meeting_dir) if meeting_dir else self.db_path.parent.parent
        self.wg = wg
        self.meeting_id = meeting_id
        self.mode = mode
        self.leaf_model_spec = leaf_model_spec
        self.synthesis_model_spec = synthesis_model_spec
        self.tdoc_meta_map = tdoc_meta_map
        self.focus_prompt = focus_prompt
        self._cancel_flag = CancellationFlag()

    def request_cancel(self) -> None:
        self._cancel_flag.cancel()
        self.requestInterruption()

    def run(self) -> None:
        try:
            store = EmailSummaryStore(self.db_path)
            orch = EmailSummaryOrchestrator(
                store,
                self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
            )
            record, err = orch.summarize_ai(
                self.ai_code,
                self.ai_label,
                self.leaf_buckets,
                mode=self.mode,
                leaf_model_spec=self.leaf_model_spec,
                synthesis_model_spec=self.synthesis_model_spec,
                tdoc_meta_map=self.tdoc_meta_map,
                focus_prompt=self.focus_prompt,
                cancel_flag=self._cancel_flag,
                on_progress=self.progress_changed.emit,
                on_chunk=self.token_received.emit,
            )
            if self._cancel_flag.is_cancelled() or self.isInterruptionRequested():
                self.summary_failed.emit("Briefing cancelled by user.")
                return
            if err or record is None:
                self.summary_failed.emit(err or "Briefing generation failed.")
                return
            self.summary_completed.emit(record)
        except Exception as exc:
            if not self._cancel_flag.is_cancelled():
                self.summary_failed.emit(str(exc))
