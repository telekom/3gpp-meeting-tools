"""Modeless two-step Send & Paste-Back handoff dialog for Microsoft 365 Copilot."""

from __future__ import annotations

import os
from pathlib import Path
import re
import subprocess
from typing import Optional

from PyQt5.QtCore import QMimeData, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ai.llm_runner import get_m365_copilot_target, set_m365_copilot_target
from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.toast import notify_info, notify_success, notify_toast
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.emails.core.ai_summary import (
    EmailSummaryStore,
    M365HandoffPacket,
    copilot_html_to_markdown,
    launch_m365_copilot,
    sanitize_copilot_markdown,
    save_external_summary_from_handoff,
    stage_handoff_markdown,
)

_RICH_HTML_TAG_RE = re.compile(r"<(?:h[1-6]|ul|ol|li|table|tr|th|td|strong|b|p|style)\b", re.IGNORECASE)


def extract_clean_markdown_from_mime(source: Optional[QMimeData]) -> str:
    """Prefer Copilot rich HTML -> GFM conversion when available, else sanitize plain text."""
    if source is None:
        return ""
    if source.hasHtml():
        html_str = str(source.html() or "")
        if _RICH_HTML_TAG_RE.search(html_str):
            converted = copilot_html_to_markdown(html_str)
            if converted:
                return converted
    return sanitize_copilot_markdown(source.text() if source.hasText() else "")


class _CopilotResponseEdit(QPlainTextEdit):
    """QPlainTextEdit that converts pasted M365 Copilot HTML/text directly into clean Markdown."""

    def insertFromMimeData(self, source: Optional[QMimeData]) -> None:
        cleaned = extract_clean_markdown_from_mime(source)
        if cleaned:
            self.insertPlainText(cleaned)
        elif source is not None:
            super().insertFromMimeData(source)

_TARGET_CHOICES = (
    ("🌐 Web (copilot.cloud.microsoft)", "web"),
    ("🪟 Windows App (ms-copilot:)", "app"),
    ("📋 Clipboard Only (No Launch)", "clipboard_only"),
)


class M365CopilotHandoffDialog(QDialog):
    """Independent modeless dialog for sending email packs to M365 Copilot and saving pasted replies."""

    summary_saved = pyqtSignal(object)  # EmailSummaryRecord

    def __init__(
        self,
        store: EmailSummaryStore,
        packet: M365HandoffPacket,
        *,
        auto_launch: bool = True,
        meeting_dir: Optional[Path] = None,
        owner: Optional[QWidget] = None,
    ) -> None:
        super().__init__(None)  # None avoids Win32 GW_OWNER pinning
        if owner is not None:
            try:
                owner.destroyed.connect(self.close)
            except Exception:
                pass

        self._store = store
        self._packet = packet
        self._meeting_dir = Path(meeting_dir) if meeting_dir else store.db_path.parent.parent

        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        self.setWindowTitle(f"Microsoft 365 Copilot — {packet.scope.title}")
        self.setMinimumSize(scale_px(560), scale_px(460))
        self.resize(scale_px(760), scale_px(600))

        self._build_ui()

        if auto_launch:
            self.prepare_and_launch()
        else:
            self.response_edit.setFocus()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(10)

        # Header + Launch Target
        hdr_row = QHBoxLayout()
        hdr_row.setSpacing(8)
        title_lbl = QLabel(f"<b>💠 {self._packet.scope.scope_type.upper()}:</b> {self._packet.scope.title}", self)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_PRIMARY};")
        hdr_row.addWidget(title_lbl, 1)

        if self._packet.mode == "incremental":
            mode_txt = f"Incremental · +{self._packet.new_message_count} new / {self._packet.message_count} msgs"
        elif self._packet.mode == "ai_briefing":
            mode_txt = f"Per-AI · {self._packet.message_count} msgs"
        else:
            mode_txt = f"Full · {self._packet.message_count} msgs"
        self.mode_pill = QLabel(mode_txt, self)
        self.mode_pill.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
        )
        hdr_row.addWidget(self.mode_pill)

        hdr_row.addWidget(QLabel("Target:", self))
        self.target_combo = QComboBox(self)
        self.target_combo.setMinimumHeight(scale_px(26))
        for label, key in _TARGET_CHOICES:
            self.target_combo.addItem(label, key)
        saved_target = get_m365_copilot_target()
        idx = self.target_combo.findData(saved_target)
        if idx >= 0:
            self.target_combo.setCurrentIndex(idx)
        self.target_combo.currentIndexChanged.connect(self._on_target_changed)
        hdr_row.addWidget(self.target_combo)
        root.addLayout(hdr_row)

        # Step 1: Prompt & Optional Markdown Upload
        step1 = QFrame(self)
        step1.setObjectName("M365Step1Card")
        step1.setStyleSheet(
            f"#M365Step1Card {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        s1_lay = QVBoxLayout(step1)
        s1_lay.setContentsMargins(10, 8, 10, 8)
        s1_lay.setSpacing(6)

        s1_title = QLabel("<b>Step 1 — Send Prepared Context to M365 Copilot</b>", step1)
        s1_title.setStyleSheet(f"color: {C.INFO_DARK}; font-size: 12px;")
        s1_lay.addWidget(s1_title)

        self.step1_info_lbl = QLabel("", step1)
        self.step1_info_lbl.setWordWrap(True)
        self.step1_info_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        s1_lay.addWidget(self.step1_info_lbl)

        s1_btns = QHBoxLayout()
        s1_btns.setSpacing(6)
        btn_h = scale_px(28)

        self.btn_copy_prompt = QPushButton("📋 Copy Prompt", step1)
        self.btn_copy_prompt.setMinimumHeight(btn_h)
        self.btn_copy_prompt.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_prompt.clicked.connect(self._on_copy_prompt)
        s1_btns.addWidget(self.btn_copy_prompt)

        self.btn_stage_or_reveal = QPushButton("📂 Reveal .md File", step1)
        self.btn_stage_or_reveal.setMinimumHeight(btn_h)
        self.btn_stage_or_reveal.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_stage_or_reveal.clicked.connect(self._on_stage_or_reveal_file)
        s1_btns.addWidget(self.btn_stage_or_reveal)

        s1_btns.addStretch()

        self.btn_launch_copilot = QPushButton("🚀 Open M365 Copilot", step1)
        self.btn_launch_copilot.setMinimumHeight(btn_h)
        self.btn_launch_copilot.setCursor(Qt.PointingHandCursor)
        self.btn_launch_copilot.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_launch_copilot.clicked.connect(self._on_launch_clicked)
        s1_btns.addWidget(self.btn_launch_copilot)

        s1_lay.addLayout(s1_btns)
        root.addWidget(step1)
        self._refresh_step1_status()

        # Step 2: Paste Back Response
        step2 = QFrame(self)
        step2.setObjectName("M365Step2Card")
        step2.setStyleSheet(
            f"#M365Step2Card {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        s2_lay = QVBoxLayout(step2)
        s2_lay.setContentsMargins(10, 8, 10, 8)
        s2_lay.setSpacing(6)

        s2_title = QLabel(
            "<b>Step 2 — Paste M365 Copilot Response Back</b> "
            f"<span style='color:{C.TEXT_MUTED}; font-size:11px;'>(tracks cutoff &amp; covered message IDs automatically)</span>",
            step2,
        )
        s2_title.setWordWrap(True)
        s2_lay.addWidget(s2_title)

        self.response_edit = _CopilotResponseEdit(step2)
        self.response_edit.setPlaceholderText(
            "Paste the response copied from Microsoft 365 Copilot here (HTML or Markdown auto-cleaned), then click 'Save & Show Inline'..."
        )
        self.response_edit.setMinimumHeight(scale_px(210))
        s2_lay.addWidget(self.response_edit, 1)

        s2_btns = QHBoxLayout()
        s2_btns.setSpacing(8)
        self.btn_paste_clipboard = QPushButton("📋 Paste from Clipboard", step2)
        self.btn_paste_clipboard.setMinimumHeight(btn_h)
        self.btn_paste_clipboard.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_paste_clipboard.clicked.connect(self._on_paste_from_clipboard)
        s2_btns.addWidget(self.btn_paste_clipboard)

        s2_btns.addStretch()

        self.btn_cancel = QPushButton("Cancel", step2)
        self.btn_cancel.setMinimumHeight(btn_h)
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cancel.clicked.connect(self.close)
        s2_btns.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("💾 Save && Show Inline", step2)
        self.btn_save.setMinimumHeight(btn_h)
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; "
            f"border: none; border-radius: 4px; padding: 4px 14px; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.ACCENT_HOVER}; }}"
        )
        self.btn_save.clicked.connect(self._on_save_clicked)
        s2_btns.addWidget(self.btn_save)

        s2_lay.addLayout(s2_btns)
        root.addWidget(step2, 1)

    def _refresh_step1_status(self) -> None:
        if self._packet.attached_file is not None:
            fname = self._packet.attached_file.name
            self.step1_info_lbl.setText(
                f"Prompt is sized for file attachment ({len(self._packet.prompt_text):,} chars). "
                f"Staged file: <b>{fname}</b> — upload or drag-and-drop it into M365 Copilot, then press Ctrl+V."
            )
            self.btn_stage_or_reveal.setText("📂 Reveal .md File")
        else:
            self.step1_info_lbl.setText(
                f"Self-contained prompt ready ({len(self._packet.clipboard_prompt):,} chars). "
                "Paste directly into M365 Copilot with Ctrl+V -> Enter (no Entra ID / Graph permission needed)."
            )
            self.btn_stage_or_reveal.setText("💾 Save Prompt as .md File")

    def _on_target_changed(self) -> None:
        target = str(self.target_combo.currentData() or "web")
        set_m365_copilot_target(target)

    def prepare_and_launch(self) -> None:
        """Copy prepared prompt to clipboard and open the chosen M365 Copilot surface."""
        QApplication.clipboard().setText(self._packet.clipboard_prompt)
        target = str(self.target_combo.currentData() or get_m365_copilot_target())
        ok, msg = launch_m365_copilot(target)
        if ok:
            notify_info(self, "Ready for M365 Copilot", f"Prompt copied to clipboard. {msg}")
        else:
            notify_toast(self, "M365 Copilot Launch", msg, variant="warning")

    def _on_copy_prompt(self) -> None:
        QApplication.clipboard().setText(self._packet.clipboard_prompt)
        notify_success(self, "Prompt Copied", "Copied prepared M365 Copilot prompt to clipboard.")

    def _on_launch_clicked(self) -> None:
        QApplication.clipboard().setText(self._packet.clipboard_prompt)
        target = str(self.target_combo.currentData() or get_m365_copilot_target())
        ok, msg = launch_m365_copilot(target)
        if not ok:
            QMessageBox.warning(self, "M365 Copilot Launch", msg)
        else:
            notify_info(self, "M365 Copilot", msg)

    def _on_stage_or_reveal_file(self) -> None:
        if self._packet.attached_file is None:
            path = stage_handoff_markdown(self._meeting_dir, self._packet.scope, self._packet.prompt_text)
            self._packet.attached_file = path
            self._refresh_step1_status()
        target_path = self._packet.attached_file
        if target_path and target_path.exists():
            if os.name == "nt":
                try:
                    subprocess.Popen(["explorer", "/select,", str(target_path.resolve())])
                    return
                except Exception:
                    pass
                if hasattr(os, "startfile"):
                    os.startfile(str(target_path.parent.resolve()))
                    return
            QApplication.clipboard().setText(str(target_path.resolve()))
            notify_info(self, "Staged File", f"Saved at {target_path}")

    def _on_paste_from_clipboard(self) -> None:
        cb = QApplication.clipboard()
        cleaned = extract_clean_markdown_from_mime(cb.mimeData()) or sanitize_copilot_markdown(cb.text())
        if not cleaned.strip():
            notify_toast(self, "Clipboard Empty", "No text found on the clipboard.", variant="warning")
            return
        self.response_edit.setPlainText(cleaned)

    def _on_save_clicked(self) -> None:
        raw = self.response_edit.toPlainText()
        rec, err = save_external_summary_from_handoff(self._store, self._packet, raw)
        if err or rec is None:
            QMessageBox.warning(self, "Empty Summary", err or "Could not save pasted summary.")
            return
        notify_success(self, "M365 Copilot Summary Saved", f"Saved summary for {rec.title}.")
        self.summary_saved.emit(rec)
        self.close()
