# --- File: src/modules/emails/ui/email_summary_dialog.py ---
import datetime
from pathlib import Path
from typing import List, Dict

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPlainTextEdit,
    QPushButton, QComboBox, QMessageBox, QFrame, QApplication, QToolTip
)

from core.ai.ollama_client import OllamaClient, load_ollama_config
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.emails.core.email_summary_worker import EmailSummaryWorker


class EmailSummaryDialog(QDialog):
    """Modeless dialog streaming an AI synthesis of a TDoc email discussion thread."""

    notes_updated = pyqtSignal(str, str, str)  # (tdoc_id, status, notes)

    def __init__(
        self,
        tdoc_id: str,
        family_tdocs: List[str],
        emails_data: List[Dict],
        save_callback=None,
        parent=None
    ):
        super().__init__(parent)
        self.tdoc_id = tdoc_id.upper()
        self.family_tdocs = family_tdocs
        self.emails_data = emails_data
        self.save_callback = save_callback

        self.setWindowTitle(f"🤖 AI Discussion Summary: {self.tdoc_id}")
        self.resize(860, 640)

        # Configure as an independent, non-modal top-level window
        self.setModal(False)
        self.setWindowFlags(
            Qt.Window
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowCloseButtonHint
        )

        self.client = OllamaClient()
        self.worker: EmailSummaryWorker = None

        self._setup_ui()
        self._refresh_models()
        self._start_summary()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # --- Top Model & Status Toolbar ---
        top_strip = QHBoxLayout()
        top_strip.setSpacing(10)

        top_strip.addWidget(QLabel("🦙 Model:"))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(180)
        top_strip.addWidget(self.model_combo)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.clicked.connect(self._refresh_models)
        top_strip.addWidget(self.refresh_btn)

        top_strip.addStretch()

        self.status_lbl = QLabel("Ready.")
        self.status_lbl.setStyleSheet("color: #64748B; font-weight: bold;")
        top_strip.addWidget(self.status_lbl)
        layout.addLayout(top_strip)

        # --- Thread Metadata Card ---
        meta_card = QFrame()
        meta_card.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        meta_layout = QVBoxLayout(meta_card)
        meta_layout.setSpacing(4)
        meta_layout.setContentsMargins(8, 6, 8, 6)

        info_line = QHBoxLayout()
        tdoc_badge = QLabel(f"<b>{self.tdoc_id}</b>")
        tdoc_badge.setStyleSheet("font-size: 14px; color: #005A9E; font-weight: bold;")
        info_line.addWidget(tdoc_badge)

        chips_text = " ➔ ".join(self.family_tdocs)
        fam_lbl = QLabel(f"Family: <b>{chips_text}</b>")
        fam_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        info_line.addWidget(fam_lbl)

        count_lbl = QLabel(f"Messages: <b>{len(self.emails_data)}</b>")
        count_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        info_line.addWidget(count_lbl)

        info_line.addStretch()
        meta_layout.addLayout(info_line)
        layout.addWidget(meta_card)

        # --- Streaming Output Area ---
        self.output_text = QPlainTextEdit()
        self.output_text.setReadOnly(False)
        self.output_text.setPlaceholderText("Streaming AI discussion synthesis...")
        self.output_text.setStyleSheet("""
            QPlainTextEdit {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                line-height: 1.5;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 10px;
            }
        """)
        layout.addWidget(self.output_text, stretch=1)

        # --- Bottom Action Bar ---
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(8)

        self.save_notes_btn = QPushButton("💾 Save to My Notes")
        self.save_notes_btn.setEnabled(False)
        self.save_notes_btn.setStyleSheet("""
            QPushButton {
                background-color: #0C6B0C; color: white; font-weight: bold;
                border-radius: 4px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #095209; }
            QPushButton:disabled { background-color: #E2E8F0; color: #94A3B8; }
        """)
        self.save_notes_btn.clicked.connect(self._save_to_notes)
        btn_bar.addWidget(self.save_notes_btn)

        self.copy_btn = QPushButton("📋 Copy")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._copy_summary)
        btn_bar.addWidget(self.copy_btn)

        btn_bar.addStretch()

        self.cancel_btn = QPushButton("⏹️ Stop")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_summary)
        btn_bar.addWidget(self.cancel_btn)

        self.regen_btn = QPushButton("🔄 Re-analyze")
        self.regen_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.regen_btn.clicked.connect(self._start_summary)
        btn_bar.addWidget(self.regen_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_btn.clicked.connect(self.close)
        btn_bar.addWidget(self.close_btn)

        layout.addLayout(btn_bar)

    def _refresh_models(self):
        is_online, models, _ = self.client.ping_and_get_models()
        self.model_combo.clear()
        if is_online and models:
            self.model_combo.addItems(models)
            cfg = load_ollama_config()
            selected = cfg.get("selected_model", "")
            if selected in models:
                self.model_combo.setCurrentText(selected)
            elif "qwen2.5:7b" in models:
                self.model_combo.setCurrentText("qwen2.5:7b")
            self.status_lbl.setText(f"🟢 Connected ({len(models)} models)")
            self.status_lbl.setStyleSheet("color: #16A34A; font-weight: bold;")
            self.regen_btn.setEnabled(True)
        else:
            self.status_lbl.setText("🔴 Ollama Offline")
            self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
            self.regen_btn.setEnabled(False)

    def _start_summary(self):
        model = self.model_combo.currentText().strip()
        if not model:
            self.status_lbl.setText("⚠️ Select an Ollama model first")
            return

        self.output_text.clear()
        self.status_lbl.setText("⏳ Initializing worker...")
        self.status_lbl.setStyleSheet("color: #2563EB; font-weight: bold;")
        self.cancel_btn.setEnabled(True)
        self.regen_btn.setEnabled(False)
        self.save_notes_btn.setEnabled(False)

        self.worker = EmailSummaryWorker(
            client=self.client,
            model=model,
            tdoc_id=self.tdoc_id,
            family_tdocs=self.family_tdocs,
            emails_data=self.emails_data,
            parent=self
        )
        self.worker.progress_status.connect(self._handle_progress)
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _handle_progress(self, msg: str):
        self.status_lbl.setText(msg)

    def _handle_token(self, delta: str):
        self.output_text.insertPlainText(delta)
        self.output_text.ensureCursorVisible()

    def _handle_complete(self, full_text: str):
        self.status_lbl.setText("✅ Completed")
        self.status_lbl.setStyleSheet("color: #16A34A; font-weight: bold;")
        self.cancel_btn.setEnabled(False)
        self.regen_btn.setEnabled(True)
        self.save_notes_btn.setEnabled(True)

    def _handle_failed(self, err_msg: str):
        self.status_lbl.setText("⚠️ Failed / Cancelled")
        self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
        self.cancel_btn.setEnabled(False)
        self.regen_btn.setEnabled(True)
        if self.output_text.toPlainText().strip():
            self.save_notes_btn.setEnabled(True)

    def _cancel_summary(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.status_lbl.setText("Cancelling...")

    def _save_to_notes(self):
        summary_text = self.output_text.toPlainText().strip()
        if not summary_text:
            QMessageBox.warning(self, "Empty Summary", "No summary text available to save.")
            return

        model_name = self.model_combo.currentText().strip() or "Ollama"
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        header = f"--- 🤖 AI Email Summary ({model_name} @ {timestamp}) ---"

        if self.save_callback:
            self.save_callback(self.tdoc_id, None, f"{header}\n{summary_text}")

        self.save_notes_btn.setText("✅ Saved to Notes!")
        self.save_notes_btn.setStyleSheet("""
            QPushButton {
                background-color: #15803D; color: white; font-weight: bold;
                border-radius: 4px; padding: 6px 14px;
            }
        """)

    def _copy_summary(self):
        text = self.output_text.toPlainText().strip()
        if text:
            QApplication.clipboard().setText(text)
            QToolTip.showText(QCursor.pos(), "📋 Copied summary to clipboard!", self)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(300)
        event.accept()