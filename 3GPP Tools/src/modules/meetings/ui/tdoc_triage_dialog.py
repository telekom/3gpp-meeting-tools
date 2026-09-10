# --- File: src/modules/meetings/ui/tdoc_triage_dialog.py ---
import datetime
from pathlib import Path
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPlainTextEdit,
    QPushButton, QComboBox, QMessageBox, QFrame, QApplication, QToolTip
)

from core.ai.ollama_client import OllamaClient, load_ollama_config
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.tdoc_triage_worker import TDocTriageWorker
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS


class TDocTriageDialog(QDialog):
    """Modeless technical triage dialog for streaming AI proposal summaries."""

    notes_updated = pyqtSignal(str, str, str)  # (tdoc_id, status, notes)

    def __init__(
        self,
        tdoc_data: dict,
        meeting_dir: Path,
        docs_ftp_url: str = "",
        revisions_url: str = "",
        save_callback=None,
        parent=None,
    ):
        super().__init__(parent)
        self.tdoc_data = tdoc_data
        self.tdoc_id = str(tdoc_data.get("TDoc", "")).strip().upper()
        self.meeting_dir = Path(meeting_dir)
        self.docs_ftp_url = docs_ftp_url
        self.revisions_url = revisions_url
        self.save_callback = save_callback

        self.setWindowTitle(f"🤖 AI Technical Triage: {self.tdoc_id}")
        self.resize(860, 640)

        # Modeless window controls
        self.setModal(False)
        self.setWindowFlags(
            Qt.Window
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowCloseButtonHint
        )

        self.client = OllamaClient()
        self.worker: TDocTriageWorker = None

        self._setup_ui()
        self._refresh_models()
        self._start_triage()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # --- Top Control Strip ---
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

        # --- Metadata Card ---
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

        title_line = QHBoxLayout()
        tdoc_badge = QLabel(f"<b>{self.tdoc_id}</b>")
        tdoc_badge.setStyleSheet("font-size: 14px; color: #005A9E; font-weight: bold;")
        title_line.addWidget(tdoc_badge)

        doc_type = str(self.tdoc_data.get("Type", "CR"))
        type_lbl = QLabel(f"[{doc_type}]")
        type_lbl.setStyleSheet("color: #64748B; font-weight: bold;")
        title_line.addWidget(type_lbl)

        ai_str = str(self.tdoc_data.get("Agenda Item", "N/A"))
        ai_lbl = QLabel(f"AI: <b>{ai_str}</b>")
        ai_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        title_line.addWidget(ai_lbl)

        src_str = str(self.tdoc_data.get("Source", "Unknown"))
        src_lbl = QLabel(f"Source: <b>{src_str}</b>")
        src_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        title_line.addWidget(src_lbl)

        title_line.addStretch()
        meta_layout.addLayout(title_line)

        title_txt = str(self.tdoc_data.get("Title", "No Title Available")).strip()
        doc_title_lbl = QLabel(title_txt)
        doc_title_lbl.setWordWrap(True)
        doc_title_lbl.setStyleSheet("color: #475569; font-size: 12px;")
        meta_layout.addWidget(doc_title_lbl)

        layout.addWidget(meta_card)

        # --- Streaming Output Area ---
        self.output_text = QPlainTextEdit()
        self.output_text.setReadOnly(False)
        self.output_text.setPlaceholderText("Streaming AI triage analysis...")
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

        # My Status Combo
        btn_bar.addWidget(QLabel("<b>My Status:</b>"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(MY_STATUS_OPTIONS)
        curr_status = str(self.tdoc_data.get("My Status", "⚪ Neutral")).replace("🔄 ", "").strip()
        self.status_combo.setCurrentText(
            curr_status if curr_status in MY_STATUS_OPTIONS else "⚪ Neutral"
        )
        self.status_combo.setStyleSheet("padding: 4px 8px; background: white; border: 1px solid #CBD5E1;")
        btn_bar.addWidget(self.status_combo)

        btn_bar.addSpacing(10)

        # Save to My Notes
        self.save_notes_btn = QPushButton("💾 Save to My Notes")
        self.save_notes_btn.setObjectName("primaryBtn")
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

        # Copy Summary
        self.copy_btn = QPushButton("📋 Copy")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._copy_summary)
        btn_bar.addWidget(self.copy_btn)

        btn_bar.addStretch()

        # Stop / Cancel
        self.cancel_btn = QPushButton("⏹️ Stop")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_triage)
        btn_bar.addWidget(self.cancel_btn)

        # Regenerate
        self.regen_btn = QPushButton("🔄 Re-analyze")
        self.regen_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.regen_btn.clicked.connect(self._start_triage)
        btn_bar.addWidget(self.regen_btn)

        # Close
        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_btn.clicked.connect(self.close)
        btn_bar.addWidget(self.close_btn)

        layout.addLayout(btn_bar)

    def _refresh_models(self):
        is_online, models, err = self.client.ping_and_get_models()
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

    def _start_triage(self):
        model = self.model_combo.currentText().strip()
        if not model:
            self.status_lbl.setText("⚠️ Select an Ollama model first")
            return

        self.output_text.clear()
        self.status_lbl.setText("⏳ Initializing triage worker...")
        self.status_lbl.setStyleSheet("color: #2563EB; font-weight: bold;")
        self.cancel_btn.setEnabled(True)
        self.regen_btn.setEnabled(False)
        self.save_notes_btn.setEnabled(False)

        self.worker = TDocTriageWorker(
            client=self.client,
            model=model,
            tdoc_data=self.tdoc_data,
            meeting_dir=self.meeting_dir,
            docs_ftp_url=self.docs_ftp_url,
            revisions_url=self.revisions_url,
            parent=self,
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

    def _cancel_triage(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.status_lbl.setText("Cancelling...")

    def _save_to_notes(self):
        """Appends the AI triage output to existing personal notes with a clear header."""
        ai_summary = self.output_text.toPlainText().strip()
        if not ai_summary:
            QMessageBox.warning(self, "Empty Analysis", "No analysis text available to save.")
            return

        selected_status = self.status_combo.currentText()
        model_name = self.model_combo.currentText().strip() or "Ollama"
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

        header = f"--- 🤖 AI Triage ({model_name} @ {timestamp}) ---"

        existing_notes = (
            str(self.tdoc_data.get("My Notes", ""))
            .replace("🔄 [From Base]: ", "")
            .replace("🔄 [From Base]", "")
            .strip()
        )

        if existing_notes:
            combined_notes = f"{existing_notes}\n\n{header}\n{ai_summary}"
        else:
            combined_notes = f"{header}\n{ai_summary}"

        # Update local dictionary reference so repeated saves don't duplicate
        self.tdoc_data["My Notes"] = combined_notes
        self.tdoc_data["My Status"] = selected_status

        if self.save_callback:
            self.save_callback(self.tdoc_id, selected_status, combined_notes)

        self.notes_updated.emit(self.tdoc_id, selected_status, combined_notes)

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