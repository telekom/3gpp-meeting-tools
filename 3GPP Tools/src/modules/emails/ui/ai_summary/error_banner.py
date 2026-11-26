"""Shared compact dismissible error banner with Retry & M365 Copilot recovery actions."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.toast import notify_info


class SummaryErrorBanner(QFrame):
    """Compact inline danger banner with Retry, M365 Copilot fallback, Copy Error, and Dismiss."""

    retry_requested = pyqtSignal()
    m365_copilot_requested = pyqtSignal()
    dismiss_requested = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("SummaryErrorBanner")
        self._full_error: str = ""
        self.setStyleSheet(
            f"#SummaryErrorBanner {{ background-color: {C.DANGER_SOFT}; "
            f"border: 1px solid {C.DANGER_BORDER_HOVER}; border-radius: 5px; }}"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 6, 8, 6)
        root.setSpacing(5)

        self.msg_lbl = QLabel("", self)
        self.msg_lbl.setWordWrap(True)
        self.msg_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.msg_lbl.setStyleSheet(f"color: {C.DANGER_DARK}; font-size: 11px; font-weight: 500; background: transparent;")
        root.addWidget(self.msg_lbl)

        btn_bar = QWidget(self)
        btn_bar.setStyleSheet(".QWidget { background: transparent; border: none; }")
        flow = FlowLayout(btn_bar, margin=0, h_spacing=6, v_spacing=4)
        btn_h = scale_px(24)

        btn_qss = (
            f"QPushButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_STRONG}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 8px; font-size: 11px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }}"
        )

        self.btn_retry = QPushButton("🔁 Retry", btn_bar)
        self.btn_retry.setMinimumHeight(btn_h)
        self.btn_retry.setCursor(Qt.PointingHandCursor)
        self.btn_retry.setStyleSheet(
            f"QPushButton {{ background-color: {C.DANGER}; color: {C.TEXT_ON_ACCENT}; "
            f"border: none; border-radius: 4px; padding: 2px 10px; font-size: 11px; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.DANGER_DARK}; }}"
        )
        self.btn_retry.clicked.connect(self.retry_requested.emit)
        flow.addWidget(self.btn_retry)

        self.btn_m365 = QPushButton("💠 Try M365 Copilot", btn_bar)
        self.btn_m365.setMinimumHeight(btn_h)
        self.btn_m365.setCursor(Qt.PointingHandCursor)
        self.btn_m365.setStyleSheet(btn_qss)
        self.btn_m365.setToolTip("Prepare this discussion for Microsoft 365 Copilot instead.")
        self.btn_m365.clicked.connect(self.m365_copilot_requested.emit)
        flow.addWidget(self.btn_m365)

        self.btn_copy_err = QPushButton("📋 Copy Error", btn_bar)
        self.btn_copy_err.setMinimumHeight(btn_h)
        self.btn_copy_err.setCursor(Qt.PointingHandCursor)
        self.btn_copy_err.setStyleSheet(btn_qss)
        self.btn_copy_err.clicked.connect(self._on_copy_error)
        flow.addWidget(self.btn_copy_err)

        self.btn_dismiss = QPushButton("✕ Dismiss", btn_bar)
        self.btn_dismiss.setMinimumHeight(btn_h)
        self.btn_dismiss.setCursor(Qt.PointingHandCursor)
        self.btn_dismiss.setStyleSheet(btn_qss)
        self.btn_dismiss.clicked.connect(self.dismiss_requested.emit)
        flow.addWidget(self.btn_dismiss)

        root.addWidget(btn_bar)
        self.setVisible(False)

    def show_error(self, message: str) -> None:
        raw = str(message or "Summary generation failed.").strip()
        self._full_error = raw
        lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
        concise = " ".join(lines)[:320]
        if len(" ".join(lines)) > 320:
            concise = concise.rstrip() + "…"
        self.msg_lbl.setText(f"⚠️ {concise}")
        self.msg_lbl.setToolTip(raw)
        self.setVisible(True)

    def clear_error(self) -> None:
        self._full_error = ""
        self.msg_lbl.setText("")
        self.msg_lbl.setToolTip("")
        self.setVisible(False)

    def _on_copy_error(self) -> None:
        if self._full_error:
            QApplication.clipboard().setText(self._full_error)
            notify_info(self, "Error Copied", "Copied full summary error details to clipboard.")
