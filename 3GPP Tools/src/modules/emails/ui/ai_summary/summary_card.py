"""Shared collapsible InlineEmailSummaryCard mounted inside thread & TDoc readers."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from PyQt5.QtCore import QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QLabel,
    QMenu,
    QPushButton,
    QSizePolicy,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ai.llm_runner import ModelSpec, resolve_effective_model_spec
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.toast import notify_info, notify_success
from modules.emails.core.ai_summary import (
    EmailSummaryRecord,
    EmailSummaryStore,
    EmailSummaryTaskWorker,
    SummaryFreshness,
    SummaryScope,
    build_leaf_handoff_packet,
)
from modules.emails.core.markdown_export import apply_email_reader_document_style, format_compact_timestamp

from .error_banner import SummaryErrorBanner
from .model_picker import populate_summary_model_menu


class InlineEmailSummaryCard(QFrame):
    """
    Zero-clutter collapsible AI Summary card mounted in a vertical splitter above
    the email timeline reader. Hidden (`0 px`) when no summary exists yet.
    """

    anchor_clicked = pyqtSignal(QUrl)
    append_to_notes_requested = pyqtSignal(str)
    summary_saved = pyqtSignal(object)  # EmailSummaryRecord
    expand_toggled = pyqtSignal(bool)
    scope_visibility_changed = pyqtSignal(bool)
    header_double_clicked = pyqtSignal()

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
        self._m365_dialog = None
        self._last_error: str = ""
        self._popped_out: bool = False
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

        hdr_bar = QWidget(self)
        hdr_bar.setStyleSheet(".QWidget { background: transparent; border: none; }")
        hdr_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        hdr_sp.setHeightForWidth(True)
        hdr_bar.setSizePolicy(hdr_sp)
        hdr = FlowLayout(hdr_bar, margin=0, h_spacing=6, v_spacing=4)
        self._hdr_bar = hdr_bar
        self._hdr_flow = hdr
        hdr_bar.mouseDoubleClickEvent = lambda ev: (self.header_double_clicked.emit(), ev.accept())
        ctrl_h = scale_px(24)

        self.btn_toggle = QPushButton("▾ 🤖 AI Summary", hdr_bar)
        self.btn_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_toggle.setMinimumHeight(ctrl_h)
        self.btn_toggle.setStyleSheet(
            f"QPushButton {{ background: transparent; border: none; color: {C.INFO_DARK}; "
            f"font-weight: bold; font-size: 12px; padding: 2px 4px; text-align: left; }} "
            f"QPushButton:hover {{ color: {C.ACCENT_PRIMARY}; }}"
        )
        self.btn_toggle.clicked.connect(self._on_toggle_expand)
        hdr.addWidget(self.btn_toggle)

        self.model_pill = QLabel("", hdr_bar)
        self.model_pill.setMinimumHeight(ctrl_h)
        self.model_pill.setStyleSheet(
            f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px;"
        )
        hdr.addWidget(self.model_pill)

        self.freshness_pill = QLabel("", hdr_bar)
        self.freshness_pill.setMinimumHeight(ctrl_h)
        self.freshness_pill.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
        )
        hdr.addWidget(self.freshness_pill)

        btn_qss = (
            f"QPushButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_STRONG}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 8px; font-size: 11px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }} "
            f"QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; }}"
        )

        self.btn_update = QPushButton("⚡ Update", hdr_bar)
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

        self.btn_rebuild = QPushButton("🔄 Rebuild", hdr_bar)
        self.btn_rebuild.setCursor(Qt.PointingHandCursor)
        self.btn_rebuild.setMinimumHeight(ctrl_h)
        self.btn_rebuild.setStyleSheet(btn_qss)
        self.btn_rebuild.setToolTip("Re-summarize the full discussion from scratch.")
        self.btn_rebuild.clicked.connect(lambda: self.trigger_summarize(mode="full"))
        hdr.addWidget(self.btn_rebuild)

        self.btn_model = QPushButton("⚙️ Model ▾", hdr_bar)
        self.btn_model.setCursor(Qt.PointingHandCursor)
        self.btn_model.setMinimumHeight(ctrl_h)
        self.btn_model.setStyleSheet(btn_qss)
        self.model_menu = QMenu(self)
        self.model_menu.aboutToShow.connect(self._populate_model_menu)
        self.btn_model.setMenu(self.model_menu)
        hdr.addWidget(self.btn_model)

        self.btn_undo = QPushButton("↩️ Undo", hdr_bar)
        self.btn_undo.setCursor(Qt.PointingHandCursor)
        self.btn_undo.setMinimumHeight(ctrl_h)
        self.btn_undo.setStyleSheet(btn_qss)
        self.btn_undo.setToolTip("Revert to the previous saved summary version.")
        self.btn_undo.clicked.connect(self._on_undo_clicked)
        self.btn_undo.setVisible(False)
        hdr.addWidget(self.btn_undo)

        self.btn_copy = QPushButton("📋 Copy", hdr_bar)
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setMinimumHeight(ctrl_h)
        self.btn_copy.setStyleSheet(btn_qss)
        self.btn_copy.clicked.connect(self._on_copy_clicked)
        hdr.addWidget(self.btn_copy)

        self.btn_to_notes = QPushButton("📝 → Notes", hdr_bar)
        self.btn_to_notes.setCursor(Qt.PointingHandCursor)
        self.btn_to_notes.setMinimumHeight(ctrl_h)
        self.btn_to_notes.setStyleSheet(btn_qss)
        self.btn_to_notes.setToolTip("Append this AI summary to your delegate notes.")
        self.btn_to_notes.clicked.connect(self._on_to_notes_clicked)
        hdr.addWidget(self.btn_to_notes)

        self.btn_stop = QPushButton("⏹ Stop", hdr_bar)
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

        root.addWidget(hdr_bar)

        self.error_banner = SummaryErrorBanner(self)
        self.error_banner.retry_requested.connect(lambda: self.trigger_summarize(mode="auto"))
        self.error_banner.m365_copilot_requested.connect(lambda: self.open_m365_copilot_handoff(mode="auto", auto_launch=True))
        self.error_banner.dismiss_requested.connect(self._on_dismiss_error)
        root.addWidget(self.error_banner)

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
        self.model_menu.addSeparator()
        act_m365_send = self.model_menu.addAction("💠 Send to M365 Copilot...")
        act_m365_send.setToolTip(
            "Prepare full or incremental Markdown prompt on clipboard, open M365 Copilot, and paste the summary back."
        )
        act_m365_send.triggered.connect(lambda _=False: self.open_m365_copilot_handoff(mode="auto", auto_launch=True))
        act_m365_paste = self.model_menu.addAction("📋 Paste M365 / External Summary...")
        act_m365_paste.setToolTip("Paste a summary generated externally in M365 Copilot and store it inline.")
        act_m365_paste.triggered.connect(lambda _=False: self.open_m365_copilot_handoff(mode="auto", auto_launch=False))

    def _on_model_selected(self, spec: ModelSpec) -> None:
        self._model_override = spec
        self._refresh_header_labels()

    def has_active_content(self) -> bool:
        return bool(self._current_record is not None or self.is_running() or self._last_error)

    def collapsed_header_height(self, available_width: int = 0) -> int:
        """Return compact single/wrapped header height when collapsed into thread-only mode."""
        m = self.layout().contentsMargins() if self.layout() is not None else None
        pad_h = (m.left() + m.right()) if m is not None else 16
        pad_v = (m.top() + m.bottom()) if m is not None else 11
        inner_w = max(scale_px(220), int(available_width) - pad_h - 4) if available_width > 0 else 0
        if inner_w > 0 and self._hdr_bar.hasHeightForWidth():
            hdr_h = self._hdr_bar.heightForWidth(inner_w)
        else:
            hdr_h = self._hdr_bar.sizeHint().height()
        return max(scale_px(38), hdr_h + pad_v + 2)

    def set_expanded(self, expanded: bool, *, emit_signal: bool = False) -> None:
        self._expanded = bool(expanded)
        self._settings.setValue(self._settings_key, self._expanded)
        self._apply_expand_state()
        if emit_signal:
            self.expand_toggled.emit(self._expanded)

    def set_popped_out(self, popped_out: bool) -> None:
        self._popped_out = bool(popped_out)
        if self._popped_out:
            self.setVisible(False)
        elif not self.has_active_content():
            self.setVisible(False)
        else:
            self.setVisible(True)
            self._render_record()
            self._refresh_header_labels()

    def clear_scope(self) -> None:
        self._scope = None
        self._messages = []
        self._current_record = None
        self._current_freshness = None
        self._last_error = ""
        self.error_banner.clear_error()
        if not self.is_running():
            self.setVisible(False)
        self.scope_visibility_changed.emit(self.has_active_content())

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
        self._last_error = ""
        self.error_banner.clear_error()

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
            self._refresh_header_labels()
            self.scope_visibility_changed.emit(False)
            return self._current_freshness

        self.setVisible(not self._popped_out)
        self._render_record()
        self._refresh_header_labels()
        self.scope_visibility_changed.emit(self.has_active_content())
        return self._current_freshness

    def current_freshness(self) -> Optional[SummaryFreshness]:
        return self._current_freshness

    def is_running(self) -> bool:
        return bool(self._worker is not None and self._worker.isRunning())

    def _on_toggle_expand(self) -> None:
        self.set_expanded(not self._expanded, emit_signal=True)

    def _apply_expand_state(self) -> None:
        has_body = self._current_record is not None or self.is_running() or (self._popped_out and not self._last_error)
        self.body_browser.setVisible(self._expanded and has_body)
        arrow = "▾" if self._expanded else "▸"
        self.btn_toggle.setText(f"{arrow} 🤖 AI Summary")

    def _render_record(self) -> None:
        self._apply_expand_state()
        if self._current_record is None:
            self.body_browser.setMarkdown("*No AI summary generated yet.*")
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
            prov_norm = str(rec.provider or "").strip().lower().replace("_", "-")
            model_norm = str(rec.model or "").strip().lower().replace("_", "-")
            if prov_norm == "m365-copilot":
                prov_badge = "💠 M365 Copilot"
            elif not rec.model or model_norm == prov_norm:
                prov_badge = rec.provider or "default"
            else:
                prov_badge = f"{rec.provider} · {rec.model}"
            if rec.reasoning_effort and prov_norm != "m365-copilot":
                prov_badge += f" · {rec.reasoning_effort}"
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

        if self._last_error:
            self.freshness_pill.setText("⚠️ Summary failed")
            self.freshness_pill.setToolTip(self._last_error)
            self.freshness_pill.setStyleSheet(
                f"background-color: {C.DANGER_SOFT}; color: {C.DANGER}; "
                f"border: 1px solid {C.DANGER_BORDER_HOVER}; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
            )
            self.btn_rebuild.setText("🔁 Retry" if not has_rec else "🔄 Rebuild")
            self.btn_update.setVisible(False)
        elif fresh is not None and has_rec:
            self.btn_rebuild.setText("🔄 Rebuild")
            cutoff_disp = format_compact_timestamp(fresh.cutoff_utc) if fresh.cutoff_utc else "last run"
            if fresh.new_count > 0:
                self.freshness_pill.setText(f"🟡 +{fresh.new_count} new since {cutoff_disp}")
                self.freshness_pill.setToolTip("")
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
                self.freshness_pill.setToolTip("")
                self.freshness_pill.setStyleSheet(
                    f"background-color: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK}; "
                    f"border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
                )
                self.btn_update.setVisible(False)
        else:
            self.btn_rebuild.setText("🔄 Rebuild")
            self.freshness_pill.setText("Ready")
            self.freshness_pill.setToolTip("")
            self.freshness_pill.setStyleSheet(
                f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; "
                f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 1px 6px; font-size: 11px;"
            )
            self.btn_update.setVisible(False)

    def trigger_summarize(self, *, mode: str = "auto", model_override: Optional[ModelSpec] = None) -> None:
        if self._store is None or self._scope is None:
            return
        if not self._messages:
            notify_info(self, "No Messages", "There are no messages in this scope to summarize.")
            return
        if self.is_running():
            return
        if model_override is not None:
            self._model_override = model_override

        self._last_error = ""
        self.error_banner.clear_error()
        self.set_expanded(True, emit_signal=True)
        self.setVisible(not self._popped_out)
        self._stream_chunks = []
        self.body_browser.setVisible(True)
        self.body_browser.setPlainText("Initializing summary...")

        spec = self._active_spec()
        self.freshness_pill.setText(f"⏳ Running ({spec.short_badge()})...")
        self.freshness_pill.setStyleSheet(
            f"background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: bold;"
        )
        self._worker = EmailSummaryTaskWorker(
            self._store.db_path, self._scope, self._messages,
            meeting_dir=self._meeting_dir, wg=self._wg, meeting_id=self._meeting_id,
            mode=mode, model_spec=spec, tdoc_meta=self._tdoc_meta, parent=self,
        )
        self._worker.progress_changed.connect(self._on_worker_progress)
        self._worker.token_received.connect(self._on_worker_token)
        self._worker.summary_completed.connect(self._on_worker_completed)
        self._worker.summary_failed.connect(self._on_worker_failed)
        self._worker.finished.connect(self._on_worker_finished)
        self._refresh_header_labels()
        self.scope_visibility_changed.emit(self.has_active_content())
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
        self._last_error = ""
        self.error_banner.clear_error()
        self._current_record = record
        if self._store is not None and self._scope is not None:
            self._current_freshness = self._store.compute_freshness(
                self._scope.scope_type,
                self._scope.scope_key,
                self._messages,
            )
        self._render_record()
        self._refresh_header_labels()
        self.scope_visibility_changed.emit(self.has_active_content())
        self.summary_saved.emit(record)

    def _on_worker_failed(self, err: str) -> None:
        self._last_error = str(err or "Summary generation failed.").strip()
        self.error_banner.show_error(self._last_error)
        if not self._expanded:
            self.set_expanded(True, emit_signal=True)
        elif self._current_record is not None:
            self._render_record()
        else:
            self._apply_expand_state()
        self._refresh_header_labels()
        self.scope_visibility_changed.emit(self.has_active_content())

    def _on_dismiss_error(self) -> None:
        self._last_error = ""
        self.error_banner.clear_error()
        if self._current_record is None and not self.is_running() and not self._popped_out:
            self.setVisible(False)
            self.scope_visibility_changed.emit(False)
            return
        self._render_record()
        self._refresh_header_labels()
        self.scope_visibility_changed.emit(self.has_active_content())

    def _on_worker_finished(self) -> None:
        self._worker = None
        self._refresh_header_labels()

    def cancel_running(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.request_cancel()
            self.freshness_pill.setText("Cancelling...")

    def open_m365_copilot_handoff(self, *, mode: str = "auto", auto_launch: bool = True):
        """Open the modeless Send & Paste-Back dialog for Microsoft 365 Copilot."""
        if self._store is None or self._scope is None:
            return None
        if not self._messages:
            notify_info(self, "No Messages", "There are no messages in this scope to send to M365 Copilot.")
            return None

        from .m365_copilot_dialog import M365CopilotHandoffDialog

        if self._m365_dialog is not None:
            try:
                self._m365_dialog.close()
            except Exception:
                pass
            self._m365_dialog = None

        packet = build_leaf_handoff_packet(
            self._store, self._scope, self._messages,
            meeting_dir=self._meeting_dir, wg=self._wg, meeting_id=self._meeting_id,
            mode=mode, tdoc_meta=self._tdoc_meta,
        )
        dlg = M365CopilotHandoffDialog(
            self._store, packet, auto_launch=auto_launch, meeting_dir=self._meeting_dir, owner=self.window()
        )
        dlg.summary_saved.connect(self._on_external_summary_saved)
        dlg.destroyed.connect(lambda _=None: setattr(self, "_m365_dialog", None))
        self._m365_dialog = dlg
        dlg.show(); dlg.raise_(); dlg.activateWindow()
        return dlg

    def _on_external_summary_saved(self, record: EmailSummaryRecord) -> None:
        self.setVisible(not self._popped_out)
        self._on_worker_completed(record)
        self.set_expanded(True, emit_signal=True)

    def shutdown_worker(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.request_cancel()
            self._worker.wait(1000)
        if self._m365_dialog is not None:
            try:
                self._m365_dialog.close()
            except Exception:
                pass
            self._m365_dialog = None

    def _on_undo_clicked(self) -> None:
        if self._store is None or self._scope is None:
            return
        restored = self._store.revert_previous_version(self._scope.scope_type, self._scope.scope_key)
        if restored is not None:
            self._current_record = restored
            self._current_freshness = self._store.compute_freshness(self._scope.scope_type, self._scope.scope_key, self._messages)
            self._render_record(); self._refresh_header_labels()
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
