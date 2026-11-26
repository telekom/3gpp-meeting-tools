"""Shared collapsible InlineEmailSummaryCard mounted inside thread & TDoc readers."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PyQt5.QtCore import QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QSizePolicy,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ai.llm_runner import ModelSpec, resolve_effective_model_spec
from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.toast import notify_info, notify_success
from modules.emails.core.ai_summary import (
    EmailSummaryRecord,
    EmailSummaryStore,
    EmailSummaryTaskWorker,
    SummaryFreshness,
    SummaryScope,
)
from modules.emails.core.markdown_export import apply_email_reader_document_style, format_compact_timestamp

from .model_picker import populate_summary_model_menu


class InlineEmailSummaryCard(QFrame):
    """
    Zero-clutter collapsible AI Summary card mounted in a vertical splitter above
    the email timeline reader. Hidden (`0 px`) when no summary exists yet.
    """

    anchor_clicked = pyqtSignal(QUrl)
    append_to_notes_requested = pyqtSignal(str)
    summary_saved = pyqtSignal(object)  # EmailSummaryRecord

    def __init__(self, parent: Optional[QWidget] = None, *, settings_key: str = "inline_summary_expanded") -> None:
        super().__init__(parent)
        self.setObjectName("InlineEmailSummaryCard")
        self._settings = QSettings("3GPP_Tools", "InlineEmailSummaryCard")
        self._settings_key = settings_key
        self._expanded = bool(self._settings.value(self._settings_key, True, type=bool))

        self._store: Optional[EmailSummaryStore] = None
        self._scope: Optional[SummaryScope] = None
        self._messages: List[Dict[str, Any]] = []
        self._meeting_dir: Optional[Path] = None
        self._wg: str = ""
        self._meeting_id: str = ""
        self._tdoc_meta: Optional[Dict[str, Any]] = None
        self._current_record: Optional[EmailSummaryRecord] = None
        self._current_freshness: Optional[SummaryFreshness] = None
        self._model_override: Optional[ModelSpec] = None
        self._worker: Optional[EmailSummaryTaskWorker] = None
        self._stream_chunks: List[str] = []

        self.setStyleSheet(
            f"#InlineEmailSummaryCard {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.INFO_BORDER}; border-radius: 6px; }}"
        )
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self._build_ui()
        self.setVisible(False)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 5, 8, 6)
        root.setSpacing(4)

        hdr = QHBoxLayout()
        hdr.setSpacing(6)
        ctrl_h = scale_px(24)

        self.btn_toggle = QPushButton("▾ 🤖 AI Summary", self)
        self.btn_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_toggle.setMinimumHeight(ctrl_h)
        self.btn_toggle.setStyleSheet(
            f"QPushButton {{ background: transparent; border: none; color: {C.INFO_DARK}; "
            f"font-weight: bold; font-size: 12px; padding: 2px 4px; text-align: left; }} "
            f"QPushButton:hover {{ color: {C.ACCENT_PRIMARY}; }}"
        )
        self.btn_toggle.clicked.connect(self._on_toggle_expand)
        hdr.addWidget(self.btn_toggle)

        self.model_pill = QLabel("", self)
        self.model_pill.setStyleSheet(
            f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px;"
        )
        hdr.addWidget(self.model_pill)

        self.freshness_pill = QLabel("", self)
        self.freshness_pill.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
        )
        hdr.addWidget(self.freshness_pill)

        hdr.addStretch()

        btn_qss = (
            f"QPushButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_STRONG}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 8px; font-size: 11px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }} "
            f"QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; }}"
        )

        self.btn_update = QPushButton("⚡ Update", self)
        self.btn_update.setCursor(Qt.PointingHandCursor)
        self.btn_update.setMinimumHeight(ctrl_h)
        self.btn_update.setStyleSheet(
            f"QPushButton {{ background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; "
            f"border: 1px solid {C.WARNING_BORDER_STRONG}; border-radius: 4px; padding: 2px 8px; "
            f"font-size: 11px; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.WARNING_BORDER}; }}"
        )
        self.btn_update.clicked.connect(lambda: self.trigger_summarize(mode="auto"))
        hdr.addWidget(self.btn_update)

        self.btn_rebuild = QPushButton("🔄 Rebuild", self)
        self.btn_rebuild.setCursor(Qt.PointingHandCursor)
        self.btn_rebuild.setMinimumHeight(ctrl_h)
        self.btn_rebuild.setStyleSheet(btn_qss)
        self.btn_rebuild.setToolTip("Re-summarize the full discussion from scratch.")
        self.btn_rebuild.clicked.connect(lambda: self.trigger_summarize(mode="full"))
        hdr.addWidget(self.btn_rebuild)

        self.btn_model = QPushButton("⚙️ Model ▾", self)
        self.btn_model.setCursor(Qt.PointingHandCursor)
        self.btn_model.setMinimumHeight(ctrl_h)
        self.btn_model.setStyleSheet(btn_qss)
        self.model_menu = QMenu(self)
        self.model_menu.aboutToShow.connect(self._populate_model_menu)
        self.btn_model.setMenu(self.model_menu)
        hdr.addWidget(self.btn_model)

        self.btn_undo = QPushButton("↩️ Undo", self)
        self.btn_undo.setCursor(Qt.PointingHandCursor)
        self.btn_undo.setMinimumHeight(ctrl_h)
        self.btn_undo.setStyleSheet(btn_qss)
        self.btn_undo.setToolTip("Revert to the previous saved summary version.")
        self.btn_undo.clicked.connect(self._on_undo_clicked)
        self.btn_undo.setVisible(False)
        hdr.addWidget(self.btn_undo)

        self.btn_copy = QPushButton("📋 Copy", self)
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setMinimumHeight(ctrl_h)
        self.btn_copy.setStyleSheet(btn_qss)
        self.btn_copy.clicked.connect(self._on_copy_clicked)
        hdr.addWidget(self.btn_copy)

        self.btn_to_notes = QPushButton("📝 → Notes", self)
        self.btn_to_notes.setCursor(Qt.PointingHandCursor)
        self.btn_to_notes.setMinimumHeight(ctrl_h)
        self.btn_to_notes.setStyleSheet(btn_qss)
        self.btn_to_notes.setToolTip("Append this AI summary to your delegate notes.")
        self.btn_to_notes.clicked.connect(self._on_to_notes_clicked)
        hdr.addWidget(self.btn_to_notes)

        self.btn_stop = QPushButton("⏹ Stop", self)
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.setMinimumHeight(ctrl_h)
        self.btn_stop.setStyleSheet(
            f"QPushButton {{ background-color: {C.DANGER_SOFT}; color: {C.DANGER}; "
            f"border: 1px solid {C.DANGER_BORDER_HOVER}; border-radius: 4px; padding: 2px 8px; "
            f"font-size: 11px; font-weight: bold; }}"
        )
        self.btn_stop.clicked.connect(self.cancel_running)
        self.btn_stop.setVisible(False)
        hdr.addWidget(self.btn_stop)

        root.addLayout(hdr)

        self.body_browser = QTextBrowser(self)
        self.body_browser.setOpenLinks(False)
        apply_email_reader_document_style(self.body_browser)
        self.body_browser.anchorClicked.connect(self.anchor_clicked.emit)
        bind_reader_zoom_settings(self.body_browser, self._settings, key="zoom_steps")
        self.body_browser.setMinimumHeight(scale_px(90))
        root.addWidget(self.body_browser, 1)

    def _active_spec(self) -> ModelSpec:
        role = "synthesis" if (self._scope and self._scope.scope_type in ("ai", "selection")) else "leaf"
        return resolve_effective_model_spec(role=role, override=self._model_override)

    def _populate_model_menu(self) -> None:
        role = "synthesis" if (self._scope and self._scope.scope_type in ("ai", "selection")) else "leaf"
        populate_summary_model_menu(
            self.model_menu,
            role=role,
            active_spec=self._active_spec(),
            on_selected=self._on_model_selected,
            on_defaults_saved=self._refresh_header_labels,
            parent=self,
        )

    def _on_model_selected(self, spec: ModelSpec) -> None:
        self._model_override = spec
        self._refresh_header_labels()

    def clear_scope(self) -> None:
        self._scope = None
        self._messages = []
        self._current_record = None
        self._current_freshness = None
        if not self.is_running():
            self.setVisible(False)

    def bind_scope(
        self,
        store: Optional[EmailSummaryStore],
        scope: SummaryScope,
        messages: Sequence[Dict[str, Any]],
        *,
        meeting_dir: Optional[Path] = None,
        wg: str = "",
        meeting_id: str = "",
        tdoc_meta: Optional[Dict[str, Any]] = None,
    ) -> Optional[SummaryFreshness]:
        self._store = store
        self._scope = scope.normalized()
        self._messages = list(messages or [])
        self._meeting_dir = meeting_dir or (store.db_path.parent.parent if store else None)
        self._wg = wg
        self._meeting_id = meeting_id
        self._tdoc_meta = tdoc_meta

        if store is None or not self._scope.scope_key:
            self.clear_scope()
            return None

        self._current_record = store.get_summary(self._scope.scope_type, self._scope.scope_key)
        self._current_freshness = store.compute_freshness(
            self._scope.scope_type,
            self._scope.scope_key,
            self._messages,
        )

        if self._current_record is None and not self.is_running():
            self.setVisible(False)
            return self._current_freshness

        self.setVisible(True)
        self._render_record()
        self._refresh_header_labels()
        return self._current_freshness

    def current_freshness(self) -> Optional[SummaryFreshness]:
        return self._current_freshness

    def is_running(self) -> bool:
        return bool(self._worker is not None and self._worker.isRunning())

    def _on_toggle_expand(self) -> None:
        self._expanded = not self._expanded
        self._settings.setValue(self._settings_key, self._expanded)
        self._apply_expand_state()

    def _apply_expand_state(self) -> None:
        self.body_browser.setVisible(self._expanded)
        arrow = "▾" if self._expanded else "▸"
        self.btn_toggle.setText(f"{arrow} 🤖 AI Summary")

    def _render_record(self) -> None:
        self._apply_expand_state()
        if self._current_record is None:
            self.body_browser.setMarkdown("*No summary generated yet.*")
            return
        self.body_browser.setMarkdown(self._current_record.summary_md)

    def _refresh_header_labels(self) -> None:
        running = self.is_running()
        self.btn_stop.setVisible(running)
        self.btn_rebuild.setEnabled(not running and bool(self._messages))
        self.btn_model.setEnabled(not running)

        rec = self._current_record
        fresh = self._current_freshness
        if rec is not None and not running:
            prov_badge = f"{rec.provider} · {rec.model or 'default'}" + (
                f" · {rec.reasoning_effort}" if rec.reasoning_effort else ""
            )
            self.model_pill.setText(prov_badge)
            self.model_pill.setToolTip(f"Generated with {prov_badge}\nActive override for next run: {self._active_spec().short_badge()}")
        else:
            self.model_pill.setText(self._active_spec().short_badge())

        if running:
            self.btn_update.setVisible(False)
            self.btn_undo.setVisible(False)
            self.btn_copy.setEnabled(False)
            self.btn_to_notes.setEnabled(False)
            return

        has_rec = rec is not None
        self.btn_copy.setEnabled(has_rec)
        self.btn_to_notes.setEnabled(has_rec)
        can_undo = bool(
            has_rec
            and self._store is not None
            and self._scope is not None
            and self._store.has_history(self._scope.scope_type, self._scope.scope_key)
        )
        self.btn_undo.setVisible(can_undo)

        if fresh is not None and has_rec:
            cutoff_disp = format_compact_timestamp(fresh.cutoff_utc) if fresh.cutoff_utc else "last run"
            if fresh.new_count > 0:
                self.freshness_pill.setText(f"🟡 +{fresh.new_count} new since {cutoff_disp}")
                self.freshness_pill.setStyleSheet(
                    f"background-color: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; "
                    f"border: 1px solid {C.WARNING_BORDER}; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
                )
                self.btn_update.setVisible(True)
                self.btn_update.setText(
                    f"⚡ Rebuild (+{fresh.new_count})" if fresh.recommend_full_rebuild else f"⚡ Update (+{fresh.new_count})"
                )
            else:
                self.freshness_pill.setText(f"✅ Up to date ({fresh.covered_count} msgs)")
                self.freshness_pill.setStyleSheet(
                    f"background-color: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK}; "
                    f"border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
                )
                self.btn_update.setVisible(False)
        else:
            self.freshness_pill.setText("Ready")
            self.btn_update.setVisible(False)

    def trigger_summarize(
        self,
        *,
        mode: str = "auto",
        model_override: Optional[ModelSpec] = None,
    ) -> None:
        if self._store is None or self._scope is None:
            return
        if not self._messages:
            notify_info(self, "No Messages", "There are no messages in this scope to summarize.")
            return
        if self.is_running():
            return
        if model_override is not None:
            self._model_override = model_override

        self._expanded = True
        self._settings.setValue(self._settings_key, True)
        self.setVisible(True)
        self._apply_expand_state()
        self._stream_chunks = []
        self.body_browser.setPlainText("Initializing summary...")

        spec = self._active_spec()
        self.freshness_pill.setText(f"⏳ Running ({spec.short_badge()})...")
        self.freshness_pill.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
        )

        self._worker = EmailSummaryTaskWorker(
            self._store.db_path,
            self._scope,
            self._messages,
            meeting_dir=self._meeting_dir,
            wg=self._wg,
            meeting_id=self._meeting_id,
            mode=mode,
            model_spec=spec,
            tdoc_meta=self._tdoc_meta,
            parent=self,
        )
        self._worker.progress_changed.connect(self._on_worker_progress)
        self._worker.token_received.connect(self._on_worker_token)
        self._worker.summary_completed.connect(self._on_worker_completed)
        self._worker.summary_failed.connect(self._on_worker_failed)
        self._worker.finished.connect(self._on_worker_finished)
        self._refresh_header_labels()
        self._worker.start()

    def _on_worker_progress(self, msg: str) -> None:
        self.freshness_pill.setText(f"⏳ {msg[:52]}")

    def _on_worker_token(self, chunk: str) -> None:
        if chunk:
            if not self._stream_chunks:
                self.body_browser.clear()
            self._stream_chunks.append(chunk)
            self.body_browser.insertPlainText(chunk)
            self.body_browser.ensureCursorVisible()

    def _on_worker_completed(self, record: EmailSummaryRecord) -> None:
        self._current_record = record
        if self._store is not None and self._scope is not None:
            self._current_freshness = self._store.compute_freshness(
                self._scope.scope_type,
                self._scope.scope_key,
                self._messages,
            )
        self._render_record()
        self._refresh_header_labels()
        self.summary_saved.emit(record)

    def _on_worker_failed(self, err: str) -> None:
        self.freshness_pill.setText(f"⚠️ {err[:58]}")
        self.freshness_pill.setStyleSheet(
            f"background-color: {C.DANGER_SOFT}; color: {C.DANGER}; border: 1px solid {C.DANGER_BORDER_HOVER}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
        )
        if self._current_record is not None:
            self._render_record()
        else:
            self.body_browser.setPlainText(f"Summary failed:\n{err}")

    def _on_worker_finished(self) -> None:
        self._worker = None
        self._refresh_header_labels()

    def cancel_running(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.request_cancel()
            self.freshness_pill.setText("Cancelling...")

    def shutdown_worker(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.request_cancel()
            self._worker.wait(1000)

    def _on_undo_clicked(self) -> None:
        if self._store is None or self._scope is None:
            return
        restored = self._store.revert_previous_version(self._scope.scope_type, self._scope.scope_key)
        if restored is not None:
            self._current_record = restored
            self._current_freshness = self._store.compute_freshness(
                self._scope.scope_type,
                self._scope.scope_key,
                self._messages,
            )
            self._render_record()
            self._refresh_header_labels()
            self.summary_saved.emit(restored)
            notify_info(self, "Summary Reverted", "Restored previous AI summary version.")

    def _on_copy_clicked(self) -> None:
        if self._current_record and self._current_record.summary_md:
            QApplication.clipboard().setText(self._current_record.summary_md)
            notify_success(self, "Summary Copied", "Copied AI summary Markdown to clipboard.")

    def _on_to_notes_clicked(self) -> None:
        if self._current_record and self._current_record.summary_md:
            tag = f"{self._current_record.provider} {self._current_record.model}".strip()
            block = f"--- 🤖 AI Summary ({tag}) ---\n{self._current_record.summary_md}"
            self.append_to_notes_requested.emit(block)
