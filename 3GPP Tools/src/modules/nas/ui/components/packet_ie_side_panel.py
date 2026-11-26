"""Collapsible inline IE Definition Side Panel embedded inside PacketStackWidget."""

from __future__ import annotations

from typing import List, Optional
from urllib.parse import unquote

from PyQt5.QtCore import QSettings, Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, COMBOBOX_STYLE_TOOLBAR
from modules.nas.core.nas_exporter import copy_definition_to_clipboard
from .packet_ie_resolver import FieldDefinitionResult, build_side_panel_html
from .packet_stack_styles import quick_btn_qss


class PacketIESidePanel(QFrame):
    """Side-by-side inline IE / clause inspector so delegates never leave Packet mode."""

    open_spec_requested = pyqtSignal(str, str, str, str)
    expand_requested = pyqtSignal(str, str, str)
    navigate_field_requested = pyqtSignal(str, str, str)
    close_requested = pyqtSignal()
    dock_toggle_requested = pyqtSignal()

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        *,
        settings: Optional[QSettings] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("packetIESidePanel")
        self.setStyleSheet(
            f"#packetIESidePanel {{ background-color: {C.BG_CARD}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
        )
        self._result: Optional[FieldDefinitionResult] = None
        self._history: List[FieldDefinitionResult] = []
        self._current_html: str = ""
        self._updating_combo: bool = False
        self._docked_bottom: bool = False
        self._setup_ui(settings)

    def _setup_ui(self, settings: Optional[QSettings]) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        hdr = QFrame(self)
        hdr.setObjectName("packetSidePanelHeader")
        hdr.setStyleSheet(
            f"#packetSidePanelHeader {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 5px; }}"
        )
        hdr_lay = QHBoxLayout(hdr)
        hdr_lay.setContentsMargins(6, 4, 6, 4)
        hdr_lay.setSpacing(4)

        btn_h = scale_px(24)
        q_qss = quick_btn_qss()

        self.back_btn = QPushButton("←", hdr)
        self.back_btn.setMinimumHeight(btn_h)
        self.back_btn.setCursor(Qt.PointingHandCursor)
        self.back_btn.setToolTip("Back to previous IE in side panel")
        self.back_btn.setStyleSheet(q_qss)
        self.back_btn.setVisible(False)
        self.back_btn.clicked.connect(self.go_back)
        hdr_lay.addWidget(self.back_btn)

        self.title_lbl = QLabel("IE Inspector", hdr)
        self.title_lbl.setMinimumWidth(scale_px(80))
        self.title_lbl.setStyleSheet(
            f"font-size: 11px; font-weight: bold; color: {C.ACCENT_SKY_DARK};"
        )
        hdr_lay.addWidget(self.title_lbl, 1)

        self.badge_lbl = QLabel("", hdr)
        self.badge_lbl.setStyleSheet(
            f"background-color: {C.SKY_SOFT}; color: {C.ACCENT_SKY_DARK}; "
            f"border: 1px solid {C.SKY_BORDER}; border-radius: 4px; "
            f"padding: 1px 6px; font-size: 10px; font-weight: 600;"
        )
        self.badge_lbl.setVisible(False)
        hdr_lay.addWidget(self.badge_lbl)

        self.version_combo = QComboBox(hdr)
        self.version_combo.setMinimumHeight(btn_h)
        self.version_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.version_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.version_combo.setToolTip("Switch indexed release for this IE definition")
        self.version_combo.setVisible(False)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)
        hdr_lay.addWidget(self.version_combo)

        self.copy_btn = QPushButton("📋 Copy", hdr)
        self.copy_btn.setMinimumHeight(btn_h)
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.setToolTip("Copy IE definition for Word / Outlook")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._on_copy_clicked)
        hdr_lay.addWidget(self.copy_btn)

        self.word_btn = QPushButton("📝 Word", hdr)
        self.word_btn.setMinimumHeight(btn_h)
        self.word_btn.setCursor(Qt.PointingHandCursor)
        self.word_btn.setToolTip("Open specification at this clause in Microsoft Word (background launch)")
        self.word_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.word_btn.clicked.connect(self._on_word_clicked)
        hdr_lay.addWidget(self.word_btn)

        self.expand_btn = QPushButton("⤢", hdr)
        self.expand_btn.setMinimumHeight(btn_h)
        self.expand_btn.setCursor(Qt.PointingHandCursor)
        self.expand_btn.setToolTip("Open in full Definition view")
        self.expand_btn.setStyleSheet(q_qss)
        self.expand_btn.clicked.connect(self._on_expand_clicked)
        hdr_lay.addWidget(self.expand_btn)

        self.dock_btn = QPushButton("⬓", hdr)
        self.dock_btn.setMinimumHeight(btn_h)
        self.dock_btn.setCursor(Qt.PointingHandCursor)
        self.dock_btn.setToolTip("Dock side panel below packet table")
        self.dock_btn.setStyleSheet(q_qss)
        self.dock_btn.clicked.connect(self.dock_toggle_requested.emit)
        hdr_lay.addWidget(self.dock_btn)

        self.close_btn = QPushButton("✕", hdr)
        self.close_btn.setMinimumHeight(btn_h)
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setToolTip("Close IE definition side panel")
        self.close_btn.setStyleSheet(q_qss)
        self.close_btn.clicked.connect(self.close_requested.emit)
        hdr_lay.addWidget(self.close_btn)

        root.addWidget(hdr)

        self.browser = QTextBrowser(self)
        self.browser.setReadOnly(True)
        self.browser.setOpenLinks(False)
        self.browser.setOpenExternalLinks(False)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setStyleSheet(
            f"QTextBrowser {{ border: 1px solid {C.BORDER_SUBTLE}; border-radius: 5px; "
            f"padding: 6px; background-color: {C.BG_CARD}; }}"
        )
        if settings is not None:
            bind_reader_zoom_settings(self.browser, settings, key="packet_stack/side_panel_zoom")
        root.addWidget(self.browser, 1)

    def set_docked_bottom(self, docked_bottom: bool) -> None:
        self._docked_bottom = bool(docked_bottom)
        if self._docked_bottom:
            self.dock_btn.setText("◫")
            self.dock_btn.setToolTip("Dock side panel to the right of packet table")
        else:
            self.dock_btn.setText("⬓")
            self.dock_btn.setToolTip("Dock side panel below packet table")

    def current_result(self) -> Optional[FieldDefinitionResult]:
        return self._result

    def show_result(self, result: FieldDefinitionResult, *, push_history: bool = True) -> None:
        if push_history and self._result is not None and self._result.title != result.title:
            self._history.append(self._result)
        self._result = result
        self._render_result(result)

    def go_back(self) -> None:
        if not self._history:
            return
        prev = self._history.pop()
        self._result = prev
        self._render_result(prev)

    def clear(self) -> None:
        self._result = None
        self._history.clear()
        self._current_html = ""
        self.back_btn.setVisible(False)
        self.version_combo.clear()
        self.version_combo.setVisible(False)
        self.badge_lbl.setVisible(False)
        self.title_lbl.setText("IE Inspector")
        self.browser.clear()

    def _render_result(self, result: FieldDefinitionResult) -> None:
        self.back_btn.setVisible(bool(self._history))
        self.title_lbl.setText(result.title)
        self.title_lbl.setToolTip(result.title)

        spec = (result.spec_number or "").strip()
        cl = (result.clause or "").strip()
        badge_parts = []
        if spec:
            badge_parts.append(f"TS {spec}" if spec[0].isdigit() else spec)
        if cl and cl != "-":
            badge_parts.append(f"Cl. {cl}")
        badge_txt = " · ".join(badge_parts)
        self.badge_lbl.setText(badge_txt)
        self.badge_lbl.setVisible(bool(badge_txt))

        self._updating_combo = True
        try:
            self.version_combo.clear()
            for idx, d in enumerate(result.defs):
                ver = str(d.get("version") or "").lstrip("vV")
                self.version_combo.addItem(f"v{ver}" if ver else f"Def #{idx + 1}", idx)
            if result.selected is not None and result.defs:
                try:
                    sel_idx = result.defs.index(result.selected)
                except ValueError:
                    sel_idx = 0
                self.version_combo.setCurrentIndex(sel_idx)
            self.version_combo.setVisible(len(result.defs) > 1)
        finally:
            self._updating_combo = False

        is_3gpp = bool(spec and spec[0].isdigit())
        self.word_btn.setEnabled(is_3gpp)
        self.expand_btn.setEnabled(result.found)

        self._current_html = build_side_panel_html(result)
        self.browser.setHtml(self._current_html)

    def _on_version_changed(self, index: int) -> None:
        if self._updating_combo or index < 0 or self._result is None:
            return
        if index < len(self._result.defs):
            self._result.selected = self._result.defs[index]
            cl = str(self._result.selected.get("clause") or self._result.clause or "").strip()
            self._result.clause = cl
            self._current_html = build_side_panel_html(self._result)
            self.browser.setHtml(self._current_html)

    def _active_spec_ver_clause(self) -> tuple[str, str, str]:
        if self._result is None:
            return ("", "", "")
        sel = self._result.selected or {}
        spec = str(sel.get("spec_number") or self._result.spec_number or "").strip()
        ver = str(sel.get("version") or "").lstrip("vV").strip()
        clause = str(sel.get("clause") or self._result.clause or "").strip()
        return (spec, ver, clause)

    def _on_copy_clicked(self) -> None:
        if self._result is None or not self._current_html:
            return
        spec, ver, clause = self._active_spec_ver_clause()
        copy_definition_to_clipboard(
            html_content=self._current_html,
            spec_number=spec,
            version=ver,
            ie_name=self._result.title,
            clause=clause,
            mode="word",
        )
        self.copy_btn.setText("✓ Copied!")
        QTimer.singleShot(1500, lambda: self.copy_btn.setText("📋 Copy"))

    def _on_word_clicked(self) -> None:
        spec, ver, clause = self._active_spec_ver_clause()
        if spec and spec[0].isdigit():
            self.open_spec_requested.emit(spec, ver, "" if clause == "-" else clause, "word")

    def _on_expand_clicked(self) -> None:
        if self._result is None:
            return
        spec, _ver, clause = self._active_spec_ver_clause()
        self.expand_requested.emit(spec, clause, self._result.title)

    def _on_anchor_clicked(self, url: QUrl) -> None:
        url_str = url.toString()
        if url_str.startswith("action:open_spec:"):
            parts = [unquote(p) for p in url_str.replace("action:open_spec:", "").split(":")]
            self.open_spec_requested.emit(
                parts[0] if len(parts) > 0 else "",
                parts[1] if len(parts) > 1 else "",
                parts[2] if len(parts) > 2 else "",
                "word",
            )
            return
        if url_str.startswith("action:packet_field:"):
            parts = [unquote(p) for p in url_str.replace("action:packet_field:", "").split(":", 2)]
            self.navigate_field_requested.emit(
                parts[0] if len(parts) > 0 else "",
                parts[1] if len(parts) > 1 else "",
                parts[2] if len(parts) > 2 else "",
            )
