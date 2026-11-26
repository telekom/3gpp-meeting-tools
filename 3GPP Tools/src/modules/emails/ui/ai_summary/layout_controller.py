"""Shared layout controller (Top/Bottom, Side-by-Side, Summary Only, Thread Only, Multi-Monitor Pop-Out)."""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtCore import QEvent, QObject, QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.emails.core.markdown_export import apply_email_reader_document_style

VALID_SUMMARY_LAYOUT_MODES = ("split_v", "split_h", "summary_only", "thread_only", "popout")


class PopoutSummaryWindow(QDialog):
    """Independent modeless top-level window mirroring the live AI Summary for multi-monitor use."""

    dock_requested = pyqtSignal()
    anchor_clicked = pyqtSignal(QUrl)
    rebuild_requested = pyqtSignal()
    copy_requested = pyqtSignal()
    to_notes_requested = pyqtSignal()

    def __init__(
        self,
        *,
        settings: QSettings,
        settings_prefix: str,
        owner: Optional[QWidget] = None,
    ) -> None:
        super().__init__(None)  # None avoids Win32 GW_OWNER pinning
        self._owner = owner
        if owner is not None:
            try:
                owner.destroyed.connect(self._on_owner_destroyed)
            except Exception:
                pass

        self._settings = settings
        self._settings_prefix = settings_prefix
        self._suppress_dock_on_close = False

        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        self.setWindowTitle("🤖 AI Email Summary")
        self.setMinimumSize(scale_px(460), scale_px(360))
        self.resize(scale_px(680), scale_px(760))
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")

        saved_geom = self._settings.value(f"{self._settings_prefix}/popout_geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(6)

        bar = QFrame(self)
        bar.setObjectName("PopoutSummaryBar")
        bar.setStyleSheet(
            f"#PopoutSummaryBar {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        bar_lay = QVBoxLayout(bar)
        bar_lay.setContentsMargins(10, 6, 10, 6)
        bar_lay.setSpacing(5)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        self.title_lbl = QLabel("<b>🤖 AI Email Summary</b>", bar)
        self.title_lbl.setWordWrap(True)
        self.title_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 13px;")
        top_row.addWidget(self.title_lbl, 1)

        self.status_lbl = QLabel("", bar)
        self.status_lbl.setStyleSheet(
            f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-radius: 4px; padding: 2px 6px; font-size: 11px;"
        )
        top_row.addWidget(self.status_lbl)
        bar_lay.addLayout(top_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        btn_h = scale_px(26)

        self.btn_rebuild = QPushButton("🔄 Rebuild", bar)
        self.btn_rebuild.setMinimumHeight(btn_h)
        self.btn_rebuild.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_rebuild.clicked.connect(self.rebuild_requested.emit)
        btn_row.addWidget(self.btn_rebuild)

        self.btn_copy = QPushButton("📋 Copy", bar)
        self.btn_copy.setMinimumHeight(btn_h)
        self.btn_copy.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy.clicked.connect(self.copy_requested.emit)
        btn_row.addWidget(self.btn_copy)

        self.btn_notes = QPushButton("📝 → Notes", bar)
        self.btn_notes.setMinimumHeight(btn_h)
        self.btn_notes.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_notes.clicked.connect(self.to_notes_requested.emit)
        btn_row.addWidget(self.btn_notes)

        btn_row.addStretch()

        self.btn_dock = QPushButton("📌 Dock Back Inline", bar)
        self.btn_dock.setMinimumHeight(btn_h)
        self.btn_dock.setCursor(Qt.PointingHandCursor)
        self.btn_dock.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_dock.setToolTip("Return the AI Summary pane back inside the email reader window.")
        self.btn_dock.clicked.connect(self._on_dock_clicked)
        btn_row.addWidget(self.btn_dock)
        bar_lay.addLayout(btn_row)
        root.addWidget(bar)

        self.browser = QTextBrowser(self)
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self.anchor_clicked.emit)
        bind_reader_zoom_settings(self.browser, self._settings, key=f"{self._settings_prefix}/popout_zoom_steps")
        root.addWidget(self.browser, 1)

    def sync_from_card(self, card) -> None:
        scope = getattr(card, "_scope", None)
        title = getattr(scope, "title", "") if scope else ""
        if title:
            self.setWindowTitle(f"🤖 AI Summary — {title}")
            self.title_lbl.setText(f"<b>🤖 AI Summary:</b> {title}")
        else:
            self.setWindowTitle("🤖 AI Email Summary")
            self.title_lbl.setText("<b>🤖 AI Email Summary</b>")

        model_txt = card.model_pill.text() if hasattr(card, "model_pill") else ""
        fresh_txt = card.freshness_pill.text() if hasattr(card, "freshness_pill") else ""
        status_bits = [b for b in (model_txt, fresh_txt) if b]
        self.status_lbl.setText(" · ".join(status_bits) or "Ready")

        rec = getattr(card, "_current_record", None)
        has_rec = rec is not None
        self.btn_copy.setEnabled(has_rec and not card.is_running())
        self.btn_notes.setEnabled(has_rec and not card.is_running())
        self.btn_rebuild.setEnabled(bool(getattr(card, "_messages", None)) and not card.is_running())

        if card.is_running():
            txt = card.body_browser.toPlainText()
            self.browser.setPlainText(txt or "Generating AI summary...")
        elif getattr(card, "_last_error", ""):
            self.browser.setMarkdown(f"### ⚠️ Summary Failed\n\n{card._last_error}" + (f"\n\n---\n\n{rec.summary_md}" if rec else ""))
        elif rec is not None:
            self.browser.setMarkdown(rec.summary_md)
        else:
            self.browser.setMarkdown("*No AI summary generated yet for this discussion.*")

    def _on_owner_destroyed(self, _obj=None) -> None:
        self._suppress_dock_on_close = True
        self._owner = None
        try:
            self.close()
        except Exception:
            pass

    def detach_owner(self) -> None:
        if self._owner is not None:
            try:
                self._owner.destroyed.disconnect(self._on_owner_destroyed)
            except Exception:
                pass
            self._owner = None

    def close_without_redock(self) -> None:
        self._suppress_dock_on_close = True
        self.detach_owner()
        self._save_geometry()
        self.close()

    def _save_geometry(self) -> None:
        try:
            self._settings.setValue(f"{self._settings_prefix}/popout_geometry", self.saveGeometry())
        except Exception:
            pass

    def _on_dock_clicked(self) -> None:
        self._save_geometry()
        self.dock_requested.emit()

    def closeEvent(self, event) -> None:
        self._save_geometry()
        super().closeEvent(event)
        if not self._suppress_dock_on_close:
            self.dock_requested.emit()


class SummaryReaderLayoutController(QObject):
    """
    Controls flexible presentation modes for an `InlineEmailSummaryCard` + email `QBrowser` pair:
    `split_v` (⬒ Top/Bottom), `split_h` (◫ Side-by-Side Summary Left | Emails Right),
    `summary_only` (⤢), `thread_only` (▸ collapsed strip), and `popout` (⧉ monitor 2 window).
    """

    mode_changed = pyqtSignal(str)

    def __init__(
        self,
        splitter: QSplitter,
        summary_card,
        reader_browser: QWidget,
        *,
        settings_prefix: str,
        owner: QWidget,
        settings: Optional[QSettings] = None,
    ) -> None:
        super().__init__(owner)
        self._splitter = splitter
        self._summary_card = summary_card
        self._reader_browser = reader_browser
        self._settings_prefix = str(settings_prefix or "email_reader").strip()
        self._owner = owner
        self._settings = settings if settings is not None else QSettings("3GPP_Tools", "EmailSummaryLayouts")
        self._popout_win: Optional[PopoutSummaryWindow] = None
        self._shutting_down = False

        saved_split = str(self._settings.value(f"{self._settings_prefix}/last_split_mode", "split_v") or "split_v")
        self._last_split_mode = saved_split if saved_split in ("split_v", "split_h") else "split_v"

        saved_mode = str(self._settings.value(f"{self._settings_prefix}/mode", self._last_split_mode) or self._last_split_mode)
        if saved_mode not in ("split_v", "split_h", "summary_only", "thread_only"):
            saved_mode = self._last_split_mode
        self._mode = saved_mode

        self._install_header_buttons()
        self._summary_card.expand_toggled.connect(self._on_card_expand_toggled)
        self._summary_card.scope_visibility_changed.connect(self._on_scope_visibility_changed)
        self._summary_card.anchor_clicked.connect(self._on_summary_anchor_clicked)
        self._summary_card.summary_saved.connect(lambda _rec: self._sync_popout())
        self._summary_card.body_browser.textChanged.connect(self._on_card_text_changed)
        self._splitter.splitterMoved.connect(self._on_splitter_moved)

        if hasattr(self._summary_card, "header_double_clicked"):
            self._summary_card.header_double_clicked.connect(self._on_header_double_clicked)

        self._apply_layout()

    def current_mode(self) -> str:
        return self._mode

    def _install_header_buttons(self) -> None:
        hdr_flow = getattr(self._summary_card, "_hdr_flow", None)
        hdr_bar = getattr(self._summary_card, "_hdr_bar", self._summary_card)
        if hdr_flow is None:
            return
        btn_h = scale_px(24)
        mode_qss = (
            f"QPushButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_STRONG}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 6px; font-size: 11px; }} "
            f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }} "
            f"QPushButton:checked {{ background-color: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
            f"border-color: {C.INFO_BORDER}; font-weight: bold; }} "
            f"QPushButton:checked:hover {{ background-color: {C.INFO_SOFT_STRONG}; }}"
        )

        self.btn_layout_v = QPushButton("⬒", hdr_bar)
        self.btn_layout_v.setCheckable(True)
        self.btn_layout_v.setMinimumHeight(btn_h)
        self.btn_layout_v.setStyleSheet(mode_qss)
        self.btn_layout_v.setToolTip("Top / Bottom Split (Summary above Emails)")
        self.btn_layout_v.clicked.connect(lambda: self.set_mode("split_v"))

        self.btn_layout_h = QPushButton("◫", hdr_bar)
        self.btn_layout_h.setCheckable(True)
        self.btn_layout_h.setMinimumHeight(btn_h)
        self.btn_layout_h.setStyleSheet(mode_qss)
        self.btn_layout_h.setToolTip("Side-by-Side Split (Summary Left | Emails Right)")
        self.btn_layout_h.clicked.connect(lambda: self.set_mode("split_h"))

        self.btn_layout_max = QPushButton("⤢", hdr_bar)
        self.btn_layout_max.setCheckable(True)
        self.btn_layout_max.setMinimumHeight(btn_h)
        self.btn_layout_max.setStyleSheet(mode_qss)
        self.btn_layout_max.setToolTip("Summary Only (Full Pane; clicking [Msg #N] restores Split)")
        self.btn_layout_max.clicked.connect(lambda: self.set_mode("summary_only"))

        self.btn_layout_popout = QPushButton("⧉", hdr_bar)
        self.btn_layout_popout.setCheckable(True)
        self.btn_layout_popout.setMinimumHeight(btn_h)
        self.btn_layout_popout.setStyleSheet(mode_qss)
        self.btn_layout_popout.setToolTip("Pop Out AI Summary to Separate Window (Multi-Monitor)")
        self.btn_layout_popout.clicked.connect(self._on_popout_clicked)

        for b in (self.btn_layout_v, self.btn_layout_h, self.btn_layout_max, self.btn_layout_popout):
            hdr_flow.addWidget(b)

    def _on_header_double_clicked(self) -> None:
        target = self._last_split_mode if self._mode in ("summary_only", "thread_only") else "summary_only"
        self.set_mode(target)

    def _on_popout_clicked(self) -> None:
        self.set_mode(self._last_split_mode if self._mode == "popout" else "popout")

    def _on_card_expand_toggled(self, expanded: bool) -> None:
        if self._mode == "popout":
            return
        self.set_mode(self._last_split_mode if expanded else "thread_only")

    def _on_scope_visibility_changed(self, _has_content: bool) -> None:
        self._sync_popout()
        self._apply_layout()

    def _on_card_text_changed(self) -> None:
        if self._popout_win is not None and self._summary_card.is_running():
            self._popout_win.sync_from_card(self._summary_card)

    def _on_summary_anchor_clicked(self, url: QUrl) -> None:
        if url.scheme().lower() == "jump" and self._mode in ("summary_only", "thread_only"):
            self.set_mode(self._last_split_mode)

    def set_mode(self, mode: str) -> None:
        clean = mode if mode in VALID_SUMMARY_LAYOUT_MODES else self._last_split_mode
        if self._mode in ("split_v", "split_h") and clean != self._mode:
            self._save_current_splitter_sizes()
        if clean in ("split_v", "split_h"):
            self._last_split_mode = clean
            self._settings.setValue(f"{self._settings_prefix}/last_split_mode", self._last_split_mode)

        was_popout = self._mode == "popout"
        self._mode = clean
        if clean != "popout":
            self._settings.setValue(f"{self._settings_prefix}/mode", clean)

        if was_popout and clean != "popout":
            self._dock_from_popout()

        if clean == "thread_only":
            self._summary_card.set_expanded(False, emit_signal=False)
        else:
            self._summary_card.set_expanded(True, emit_signal=False)

        self._apply_layout()
        self.mode_changed.emit(self._mode)

    def _sync_popout(self) -> None:
        if self._popout_win is not None:
            self._popout_win.sync_from_card(self._summary_card)

    def _dock_from_popout(self) -> None:
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            pop.close_without_redock()
        self._summary_card.set_popped_out(False)

    def _ensure_popout_window(self) -> PopoutSummaryWindow:
        if self._popout_win is None:
            pop = PopoutSummaryWindow(
                settings=self._settings,
                settings_prefix=self._settings_prefix,
                owner=self._owner,
            )
            pop.dock_requested.connect(self._on_popout_dock_requested)
            pop.anchor_clicked.connect(self._summary_card.anchor_clicked.emit)
            pop.rebuild_requested.connect(lambda: self._summary_card.trigger_summarize(mode="full"))
            pop.copy_requested.connect(self._summary_card._on_copy_clicked)
            pop.to_notes_requested.connect(self._summary_card._on_to_notes_clicked)
            self._popout_win = pop
        self._summary_card.set_popped_out(True)
        self._popout_win.sync_from_card(self._summary_card)
        self._popout_win.show()
        self._popout_win.raise_()
        self._popout_win.activateWindow()
        return self._popout_win

    def _on_popout_dock_requested(self) -> None:
        if self._shutting_down:
            return
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            pop.detach_owner()
        self._summary_card.set_popped_out(False)
        if self._mode == "popout":
            self._mode = self._last_split_mode
            self._settings.setValue(f"{self._settings_prefix}/mode", self._mode)
            self._summary_card.set_expanded(True, emit_signal=False)
            self._apply_layout()
            self.mode_changed.emit(self._mode)

    def _sync_button_checks(self) -> None:
        checks = (
            (self.btn_layout_v, self._mode == "split_v"),
            (self.btn_layout_h, self._mode == "split_h"),
            (self.btn_layout_max, self._mode == "summary_only"),
            (self.btn_layout_popout, self._mode == "popout"),
        )
        for btn, state in checks:
            btn.blockSignals(True)
            btn.setChecked(state)
            btn.blockSignals(False)

    def _save_current_splitter_sizes(self) -> None:
        if not self._summary_card.isVisible() or not self._reader_browser.isVisible():
            return
        if self._mode not in ("split_v", "split_h"):
            return
        sizes = self._splitter.sizes()
        if len(sizes) >= 2 and all(s > 20 for s in sizes[:2]):
            key = "sizes_h" if self._mode == "split_h" else "sizes_v"
            self._settings.setValue(f"{self._settings_prefix}/{key}", f"{sizes[0]},{sizes[1]}")

    def _on_splitter_moved(self, _pos: int, _index: int) -> None:
        self._save_current_splitter_sizes()

    def _restore_split_sizes(self, mode: str) -> List[int]:
        key = "sizes_h" if mode == "split_h" else "sizes_v"
        raw = str(self._settings.value(f"{self._settings_prefix}/{key}", "") or "")
        parts = [int(x) for x in raw.split(",") if x.strip().isdigit()]
        if len(parts) == 2 and parts[0] > 40 and parts[1] > 40:
            return parts
        if mode == "split_h":
            return [scale_px(430), scale_px(560)]
        return [scale_px(210), scale_px(430)]

    def _apply_layout(self) -> None:
        self._sync_button_checks()
        has_inline_content = self._summary_card.has_active_content()

        if self._mode == "popout":
            self._reader_browser.setVisible(True)
            self._ensure_popout_window()
            return

        if not has_inline_content:
            self._summary_card.setMaximumHeight(16777215)
            self._reader_browser.setVisible(True)
            return

        if self._mode == "thread_only":
            self._splitter.setOrientation(Qt.Vertical)
            collapsed_h = scale_px(38)
            self._summary_card.setMaximumHeight(collapsed_h)
            self._reader_browser.setVisible(True)
            self._splitter.setSizes([collapsed_h, scale_px(580)])
        elif self._mode == "summary_only":
            self._splitter.setOrientation(Qt.Vertical)
            self._summary_card.setMaximumHeight(16777215)
            self._reader_browser.setVisible(False)
            self._splitter.setSizes([scale_px(640), 0])
        elif self._mode == "split_h":
            self._splitter.setOrientation(Qt.Horizontal)
            self._summary_card.setMaximumHeight(16777215)
            self._reader_browser.setVisible(True)
            self._splitter.setSizes(self._restore_split_sizes("split_h"))
        else:
            self._splitter.setOrientation(Qt.Vertical)
            self._summary_card.setMaximumHeight(16777215)
            self._reader_browser.setVisible(True)
            self._splitter.setSizes(self._restore_split_sizes("split_v"))

    def shutdown(self) -> None:
        """Persist current split ratios and close any popped-out window cleanly on dialog exit."""
        self._shutting_down = True
        self._save_current_splitter_sizes()
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            self._summary_card.set_popped_out(False)
            pop.close_without_redock()
