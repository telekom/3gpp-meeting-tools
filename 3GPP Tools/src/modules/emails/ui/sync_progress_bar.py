# --- File: src/modules/emails/ui/sync_progress_bar.py ---
"""Slim, non-blocking progress and error status strip for async Outlook email sync."""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QWidget


class AsyncSyncProgressBar(QFrame):
    """
    Inline status strip displaying what Outlook sync step is running now,
    what items remain in the queue, a compact progress bar, and Cancel/Dismiss actions.
    """

    cancel_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("AsyncSyncProgressBar")
        self._apply_style(is_error=False)
        self._build_ui()
        self.setVisible(False)

    def _apply_style(self, is_error: bool):
        if is_error:
            self.setStyleSheet(
                "#AsyncSyncProgressBar { background: #FEF2F2; border: 1px solid #FECACA; border-radius: 6px; }"
            )
        else:
            self.setStyleSheet(
                "#AsyncSyncProgressBar { background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 6px; }"
            )

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(10)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("font-size: 11px; font-weight: 600; color: #1E3A8A; border: none; background: transparent;")
        layout.addWidget(self.status_label, 0)

        self.remaining_label = QLabel("")
        self.remaining_label.setStyleSheet("font-size: 11px; color: #475569; border: none; background: transparent;")
        layout.addWidget(self.remaining_label, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(130)
        self.progress_bar.setFixedHeight(10)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(
            "QProgressBar { background: #DBEAFE; border: none; border-radius: 5px; }"
            "QProgressBar::chunk { background: #2563EB; border-radius: 5px; }"
        )
        layout.addWidget(self.progress_bar)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_btn.setMinimumHeight(24)
        self.cancel_btn.setStyleSheet(
            "QPushButton { background: #FFFFFF; color: #334155; border: 1px solid #CBD5E1; "
            "border-radius: 4px; padding: 2px 10px; font-size: 11px; font-weight: 500; }"
            "QPushButton:hover { background: #F1F5F9; border-color: #94A3B8; color: #0F172A; }"
            "QPushButton:disabled { color: #94A3B8; background: #F8FAFC; }"
        )
        self.cancel_btn.clicked.connect(self._on_action_clicked)
        layout.addWidget(self.cancel_btn)

        self._error_mode = False

    def _on_action_clicked(self):
        if self._error_mode:
            self.setVisible(False)
        else:
            self.cancel_btn.setEnabled(False)
            self.cancel_btn.setText("Cancelling...")
            self.cancel_requested.emit()

    @staticmethod
    def _clip(text: str, limit: int = 40) -> str:
        t = str(text or "").strip()
        return (t[: limit - 1] + "…") if len(t) > limit else t

    def start_busy(self, title: str, total: int = 0, remaining: Optional[List[str]] = None):
        """Shows the strip in active progress mode."""
        self._error_mode = False
        self._apply_style(is_error=False)
        self.status_label.setStyleSheet("font-size: 11px; font-weight: 600; color: #1E3A8A; border: none; background: transparent;")
        self.status_label.setText(f"⏳ {title}")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, max(0, int(total)))
        self.progress_bar.setValue(0)
        self.cancel_btn.setText("Cancel")
        self.cancel_btn.setEnabled(True)
        self._set_remaining(remaining or [])
        self.setVisible(True)

    def update_progress(self, completed: int, total: int, current_item: str, remaining: List[str]):
        """Updates current item, remaining queue preview, and progress bar."""
        self._error_mode = False
        cur_step = min(completed + 1, total) if total > 0 and completed < total else completed
        self.status_label.setText(
            f"⏳ ({cur_step}/{total}) {self._clip(current_item, 44)}" if total > 0 else f"⏳ {self._clip(current_item, 50)}"
        )
        self.progress_bar.setRange(0, max(1, int(total)) if total > 0 else 0)
        self.progress_bar.setValue(min(int(completed), max(1, int(total))))
        self._set_remaining(remaining)
        self.setVisible(True)

    def _set_remaining(self, remaining: List[str]):
        rem = [str(x).strip() for x in (remaining or []) if str(x).strip()]
        if not rem:
            self.remaining_label.setText("Queue empty — finishing...")
            self.remaining_label.setToolTip("")
            return
        preview = ", ".join(self._clip(x, 22) for x in rem[:2])
        more = f" (+{len(rem) - 2} more)" if len(rem) > 2 else ""
        self.remaining_label.setText(f"Next ({len(rem)}): {preview}{more}")
        self.remaining_label.setToolTip("Remaining in queue:\n• " + "\n• ".join(rem))

    def show_error(self, message: str):
        """Switches the strip into a visible non-blocking warning/error state until dismissed."""
        self._error_mode = True
        self._apply_style(is_error=True)
        self.status_label.setStyleSheet("font-size: 11px; font-weight: 600; color: #B91C1C; border: none; background: transparent;")
        self.status_label.setText(f"⚠️ {self._clip(message, 85)}")
        self.status_label.setToolTip(str(message or ""))
        self.remaining_label.setText("")
        self.progress_bar.setVisible(False)
        self.cancel_btn.setText("Dismiss")
        self.cancel_btn.setEnabled(True)
        self.setVisible(True)

    def finish_ok(self):
        """Hides the strip cleanly when async work completes without error."""
        self._error_mode = False
        self.setVisible(False)
