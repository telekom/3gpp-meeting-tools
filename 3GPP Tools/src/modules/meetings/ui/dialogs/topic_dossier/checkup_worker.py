"""Background QThread worker for non-blocking Topic Dossier corpus checkup."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.exporters.topic_dossier.checkup import check_and_refresh_dossier
from modules.meetings.core.topic_manager import MetaTopicDefinition


class TopicCheckupWorker(QThread):
    """Runs `check_and_refresh_dossier` off the GUI thread with live progress and cancellation."""

    progress_step = pyqtSignal(int, int)
    progress_msg = pyqtSignal(str)
    checkup_finished = pyqtSignal(dict)

    def __init__(
        self,
        target_dir: Path,
        topic: Optional[MetaTopicDefinition] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.target_dir = Path(target_dir)
        self.topic = topic
        self._is_cancelled = False

    def cancel(self) -> None:
        self._is_cancelled = True
        self.requestInterruption()

    def _should_cancel(self) -> bool:
        return self._is_cancelled or self.isInterruptionRequested()

    def run(self) -> None:
        def _on_progress(cur: int, total: int, msg: str) -> None:
            if not self._should_cancel():
                self.progress_step.emit(int(cur), int(total))
                self.progress_msg.emit(str(msg))

        try:
            res = check_and_refresh_dossier(
                self.target_dir,
                topic=self.topic,
                progress_cb=_on_progress,
                cancel_check=self._should_cancel,
            )
        except Exception as exc:
            res = {
                "ok": False,
                "target_dir": str(self.target_dir),
                "error": str(exc),
            }
        self.checkup_finished.emit(res)
