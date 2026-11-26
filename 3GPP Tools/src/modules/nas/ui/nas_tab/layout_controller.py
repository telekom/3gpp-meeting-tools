"""Flexible multi-pane focus & layout controller for NASTab (Protocols workspace)."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QObject, QSettings, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

VALID_RIGHT_LAYOUT_MODES = ("split_v", "split_h", "matrix_only", "inspector_only", "popout")


def _layout_btn_qss() -> str:
    return (
        f"QPushButton {{ background-color: {C.BG_CARD}; color: {C.TEXT_STRONG}; "
        f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 6px; font-size: 11px; }} "
        f"QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }} "
        f"QPushButton:checked {{ background-color: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
        f"border-color: {C.INFO_BORDER}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.INFO_SOFT_STRONG}; }}"
    )


class PopoutInspectorWindow(QDialog):
    """Independent modeless top-level window hosting the live NASInspectorWidget."""

    dock_requested = pyqtSignal()

    def __init__(self, *, settings: QSettings, owner: Optional[QWidget] = None) -> None:
        super().__init__(None)  # None avoids Win32 GW_OWNER pinning
        self._owner = owner
        if owner is not None:
            try:
                owner.destroyed.connect(self._on_owner_destroyed)
            except Exception:
                pass

        self._settings = settings
        self._suppress_dock_on_close = False

        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setSizeGripEnabled(True)
        self.setWindowTitle("🔍 Protocol Field & Full Packet Inspector")
        self.setMinimumSize(scale_px(520), scale_px(420))
        self.resize(scale_px(760), scale_px(780))
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")

        saved_geom = self._settings.value("popout_geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        bar = QFrame(self)
        bar.setObjectName("PopoutInspectorBar")
        bar.setStyleSheet(
            f"#PopoutInspectorBar {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        bar_lay = QHBoxLayout(bar)
        bar_lay.setContentsMargins(10, 5, 10, 5)
        bar_lay.setSpacing(8)
        lbl = QLabel("<b>🔍 Protocol Inspector (Detached)</b>", bar)
        lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 12px;")
        bar_lay.addWidget(lbl, 1)

        self.btn_dock = QPushButton("📌 Dock Back Inline", bar)
        self.btn_dock.setMinimumHeight(scale_px(25))
        self.btn_dock.setCursor(Qt.PointingHandCursor)
        self.btn_dock.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_dock.setToolTip("Return the Inspector back inside the Protocols tab")
        self.btn_dock.clicked.connect(self._on_dock_clicked)
        bar_lay.addWidget(self.btn_dock)
        root.addWidget(bar)

        self.body_container = QWidget(self)
        self.body_layout = QVBoxLayout(self.body_container)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(0)
        root.addWidget(self.body_container, 1)

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
            self._settings.setValue("popout_geometry", self.saveGeometry())
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


class ProtocolWorkspaceLayoutController(QObject):
    """
    Controls Email-style flexible focus & layout ergonomics for `NASTab`:
      1. Right-pane modes (`split_v` ⬒, `split_h` ◫, `matrix_only` ⤢, `inspector_only` ⤢, `popout` ⧉)
      2. Smart auto-compact when browsing single-row `All IEs / Types` matrices
      3. Collapsible Left Sidebar (`◀`/`▶`, `Ctrl+[`)
      4. Collapsible `Indexed Protocols` card (`▾`/`▸` -> 1-line summary bar)
    """

    mode_changed = pyqtSignal(str)

    def __init__(
        self,
        *,
        owner: QWidget,
        main_splitter: QSplitter,
        left_pane: QWidget,
        left_splitter: QSplitter,
        ver_card: QFrame,
        ver_body_widget: QWidget,
        ver_summary_bar: QWidget,
        ver_summary_lbl: QLabel,
        btn_toggle_specs_card: QPushButton,
        right_splitter: QSplitter,
        matrix_card: QFrame,
        matrix_header_layout: QHBoxLayout,
        matrix_stack: QWidget,
        msg_context_banner: QFrame,
        inspector: Any,
        settings: Optional[QSettings] = None,
    ) -> None:
        super().__init__(owner)
        self._owner = owner
        self._main_splitter = main_splitter
        self._left_pane = left_pane
        self._left_splitter = left_splitter
        self._ver_card = ver_card
        self._ver_body_widget = ver_body_widget
        self._ver_summary_bar = ver_summary_bar
        self._ver_summary_lbl = ver_summary_lbl
        self._btn_toggle_specs_card = btn_toggle_specs_card
        self._right_splitter = right_splitter
        self._matrix_card = matrix_card
        self._matrix_header_layout = matrix_header_layout
        self._matrix_stack = matrix_stack
        self._msg_context_banner = msg_context_banner
        self._inspector = inspector
        self._settings = settings if settings is not None else QSettings("3GPP_Tools", "ProtocolsWorkspaceLayout")

        self._popout_win: Optional[PopoutInspectorWindow] = None
        self._shutting_down = False
        self._auto_compacted = False
        self._has_banner_content = False

        saved_split = str(self._settings.value("last_split_mode", "split_v") or "split_v")
        self._last_split_mode = saved_split if saved_split in ("split_v", "split_h") else "split_v"

        saved_mode = str(self._settings.value("right_mode", self._last_split_mode) or self._last_split_mode)
        if saved_mode not in ("split_v", "split_h", "matrix_only", "inspector_only"):
            saved_mode = self._last_split_mode
        self._right_mode = saved_mode

        self._left_collapsed = str(self._settings.value("left_collapsed", "false")).lower() == "true"
        self._specs_collapsed = str(self._settings.value("specs_collapsed", "false")).lower() == "true"

        self._install_buttons()
        self._right_splitter.splitterMoved.connect(self._on_right_splitter_moved)
        self._main_splitter.splitterMoved.connect(self._on_main_splitter_moved)
        self._left_splitter.splitterMoved.connect(self._on_left_splitter_moved)
        if hasattr(self._inspector, "header_double_clicked"):
            self._inspector.header_double_clicked.connect(self._on_inspector_double_clicked)

        self.set_specs_card_collapsed(self._specs_collapsed, persist=False)
        self.set_left_collapsed(self._left_collapsed, persist=False)
        self._apply_right_layout()

    def current_mode(self) -> str:
        return self._right_mode

    def is_left_collapsed(self) -> bool:
        return self._left_collapsed

    def is_specs_card_collapsed(self) -> bool:
        return self._specs_collapsed

    def _install_buttons(self) -> None:
        qss = _layout_btn_qss()
        btn_h = scale_px(24)

        self.btn_toggle_left = QPushButton("◀", self._matrix_card)
        self.btn_toggle_left.setCheckable(True)
        self.btn_toggle_left.setMinimumHeight(btn_h)
        self.btn_toggle_left.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_left.setStyleSheet(qss)
        self.btn_toggle_left.setToolTip("Collapse / Expand Left Protocol Sidebar (Ctrl+[)")
        self.btn_toggle_left.clicked.connect(self.toggle_left_pane)
        self._matrix_header_layout.insertWidget(0, self.btn_toggle_left)

        self.btn_layout_v = QPushButton("⬒", self._matrix_card)
        self.btn_layout_v.setCheckable(True)
        self.btn_layout_v.setMinimumHeight(btn_h)
        self.btn_layout_v.setCursor(Qt.PointingHandCursor)
        self.btn_layout_v.setStyleSheet(qss)
        self.btn_layout_v.setToolTip("Top / Bottom Split (Matrix above Inspector)")
        self.btn_layout_v.clicked.connect(lambda: self.set_mode("split_v"))
        self._matrix_header_layout.addWidget(self.btn_layout_v)

        self.btn_layout_h = QPushButton("◫", self._matrix_card)
        self.btn_layout_h.setCheckable(True)
        self.btn_layout_h.setMinimumHeight(btn_h)
        self.btn_layout_h.setCursor(Qt.PointingHandCursor)
        self.btn_layout_h.setStyleSheet(qss)
        self.btn_layout_h.setToolTip("Side-by-Side Split (Matrix Left | Inspector Right, Ctrl+L)")
        self.btn_layout_h.clicked.connect(lambda: self.set_mode("split_h"))
        self._matrix_header_layout.addWidget(self.btn_layout_h)

        self.btn_layout_matrix_max = QPushButton("⤢", self._matrix_card)
        self.btn_layout_matrix_max.setCheckable(True)
        self.btn_layout_matrix_max.setMinimumHeight(btn_h)
        self.btn_layout_matrix_max.setCursor(Qt.PointingHandCursor)
        self.btn_layout_matrix_max.setStyleSheet(qss)
        self.btn_layout_matrix_max.setToolTip("Maximize Evolution Matrix (Collapse Inspector to Strip)")
        self.btn_layout_matrix_max.clicked.connect(
            lambda: self.set_mode(self._last_split_mode if self._right_mode == "matrix_only" else "matrix_only")
        )
        self._matrix_header_layout.addWidget(self.btn_layout_matrix_max)

        insp_hdr = getattr(self._inspector, "insp_header_layout", None)
        if insp_hdr is not None:
            self.btn_layout_insp_max = QPushButton("⤢", self._inspector)
            self.btn_layout_insp_max.setCheckable(True)
            self.btn_layout_insp_max.setMinimumHeight(btn_h)
            self.btn_layout_insp_max.setCursor(Qt.PointingHandCursor)
            self.btn_layout_insp_max.setStyleSheet(qss)
            self.btn_layout_insp_max.setToolTip("Maximize Inspector / Full Packet View (Collapse Matrix to Strip)")
            self.btn_layout_insp_max.clicked.connect(self._on_inspector_double_clicked)
            insp_hdr.addWidget(self.btn_layout_insp_max)

            self.btn_layout_popout = QPushButton("⧉", self._inspector)
            self.btn_layout_popout.setCheckable(True)
            self.btn_layout_popout.setMinimumHeight(btn_h)
            self.btn_layout_popout.setCursor(Qt.PointingHandCursor)
            self.btn_layout_popout.setStyleSheet(qss)
            self.btn_layout_popout.setToolTip("Pop Out Inspector to Separate Window (Multi-Monitor)")
            self.btn_layout_popout.clicked.connect(
                lambda: self.set_mode(self._last_split_mode if self._right_mode == "popout" else "popout")
            )
            insp_hdr.addWidget(self.btn_layout_popout)
        else:
            self.btn_layout_insp_max = None
            self.btn_layout_popout = None

        self._btn_toggle_specs_card.clicked.connect(self.toggle_specs_card)

    def toggle_split_orientation(self) -> None:
        self.set_mode("split_h" if self._right_mode != "split_h" else "split_v")

    def toggle_left_pane(self) -> None:
        self.set_left_collapsed(not self._left_collapsed)

    def set_left_collapsed(self, collapsed: bool, *, persist: bool = True) -> None:
        if not collapsed and self._left_collapsed:
            pass
        elif collapsed and not self._left_collapsed:
            sizes = self._main_splitter.sizes()
            if len(sizes) >= 2 and sizes[0] >= scale_px(220):
                self._settings.setValue("left_width", sizes[0])
        self._left_collapsed = bool(collapsed)
        if persist:
            self._settings.setValue("left_collapsed", "true" if self._left_collapsed else "false")

        self._left_pane.setVisible(not self._left_collapsed)
        if not self._left_collapsed:
            saved_w = int(self._settings.value("left_width", scale_px(380)) or scale_px(380))
            saved_w = max(scale_px(280), saved_w)
            right_w = max(scale_px(420), self._main_splitter.width() - saved_w)
            self._main_splitter.setSizes([saved_w, right_w])

        self.btn_toggle_left.blockSignals(True)
        self.btn_toggle_left.setText("▶" if self._left_collapsed else "◀")
        self.btn_toggle_left.setChecked(self._left_collapsed)
        self.btn_toggle_left.blockSignals(False)

    def toggle_specs_card(self) -> None:
        self.set_specs_card_collapsed(not self._specs_collapsed)

    def set_specs_card_collapsed(self, collapsed: bool, *, persist: bool = True) -> None:
        if collapsed and not self._specs_collapsed:
            sizes = self._left_splitter.sizes()
            if len(sizes) >= 2 and sizes[0] >= scale_px(120):
                self._settings.setValue("left_splitter_sizes", f"{sizes[0]},{sizes[1]}")

        self._specs_collapsed = bool(collapsed)
        if persist:
            self._settings.setValue("specs_collapsed", "true" if self._specs_collapsed else "false")

        self._ver_body_widget.setVisible(not self._specs_collapsed)
        self._ver_summary_bar.setVisible(self._specs_collapsed)
        self._btn_toggle_specs_card.setText("▸" if self._specs_collapsed else "▾")
        if self._specs_collapsed:
            compact_h = scale_px(76)
            self._ver_card.setMaximumHeight(compact_h)
            lib_h = max(scale_px(240), self._left_splitter.height() - compact_h)
            self._left_splitter.setSizes([compact_h, lib_h])
        else:
            self._ver_card.setMaximumHeight(16777215)
            raw = str(self._settings.value("left_splitter_sizes", "") or "")
            parts = [int(x) for x in raw.split(",") if x.strip().isdigit()]
            if len(parts) == 2 and parts[0] >= scale_px(120) and parts[1] >= scale_px(140):
                self._left_splitter.setSizes(parts)
            else:
                self._left_splitter.setSizes([scale_px(270), scale_px(430)])

    def update_specs_summary(self, checked_info: List[Dict[str, Any]]) -> None:
        if not checked_info:
            self._ver_summary_lbl.setText("No specs checked — click Change to pick releases")
            return
        by_spec: Dict[str, List[str]] = {}
        for v in checked_info:
            s = str(v.get("spec_number", "")).strip()
            ver = str(v.get("version", "")).strip()
            if s:
                rel = ver.split(".")[0] if ver and ver.split(".")[0].isdigit() else ""
                tag = f"R{rel}" if rel else ver
                by_spec.setdefault(s, [])
                if tag and tag not in by_spec[s]:
                    by_spec[s].append(tag)
        chips = [f"TS {s} ({','.join(rels[:2])})" if rels else f"TS {s}" for s, rels in sorted(by_spec.items())]
        preview = " · ".join(chips[:3])
        if len(chips) > 3:
            preview += f" (+{len(chips) - 3})"
        self._ver_summary_lbl.setText(f"{len(by_spec)} active: {preview}")
        self._ver_summary_lbl.setToolTip("\n".join(chips))

    def set_banner_active(self, active: bool) -> None:
        self._has_banner_content = bool(active)
        if self._right_mode == "inspector_only":
            self._msg_context_banner.setVisible(False)

    def apply_auto_compact_for_matrix(self, browse_mode: str, row_count: int) -> None:
        """Give ~80% height to Inspector when `All IEs / Types` shows a single summary row."""
        if self._right_mode != "split_v":
            self._auto_compacted = False
            return
        if browse_mode == "ies" and 0 < row_count <= 1:
            if not self._auto_compacted:
                self._save_right_splitter_sizes()
            self._auto_compacted = True
            compact_h = scale_px(110)
            insp_h = max(scale_px(360), self._right_splitter.height() - compact_h, scale_px(480))
            self._right_splitter.setSizes([compact_h, insp_h])
        elif self._auto_compacted:
            self._auto_compacted = False
            self._right_splitter.setSizes(self._restore_right_sizes("split_v"))

    def set_mode(self, mode: str) -> None:
        clean = mode if mode in VALID_RIGHT_LAYOUT_MODES else self._last_split_mode
        if self._right_mode in ("split_v", "split_h") and clean != self._right_mode and not self._auto_compacted:
            self._save_right_splitter_sizes()
        if clean in ("split_v", "split_h"):
            self._last_split_mode = clean
            self._settings.setValue("last_split_mode", self._last_split_mode)
        self._auto_compacted = False

        was_popout = self._right_mode == "popout"
        self._right_mode = clean
        if clean != "popout":
            self._settings.setValue("right_mode", clean)

        if was_popout and clean != "popout":
            self._dock_from_popout()

        self._apply_right_layout()
        self.mode_changed.emit(self._right_mode)

    def _on_inspector_double_clicked(self) -> None:
        self.set_mode(self._last_split_mode if self._right_mode == "inspector_only" else "inspector_only")

    def _ensure_popout_window(self) -> None:
        if self._popout_win is None:
            pop = PopoutInspectorWindow(settings=self._settings, owner=self._owner)
            pop.dock_requested.connect(self._on_popout_dock_requested)
            self._popout_win = pop
        if self._inspector.parent() is not self._popout_win.body_container:
            self._popout_win.body_layout.addWidget(self._inspector)
        self._inspector.setVisible(True)
        self._inspector.stack.setVisible(True)
        self._inspector.setMaximumHeight(16777215)
        self._popout_win.show()
        self._popout_win.raise_()
        self._popout_win.activateWindow()

    def _dock_from_popout(self) -> None:
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            if self._inspector.parent() is pop.body_container:
                self._right_splitter.addWidget(self._inspector)
            pop.close_without_redock()

    def _on_popout_dock_requested(self) -> None:
        if self._shutting_down:
            return
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            if self._inspector.parent() is pop.body_container:
                self._right_splitter.addWidget(self._inspector)
            pop.detach_owner()
        if self._right_mode == "popout":
            self._right_mode = self._last_split_mode
            self._settings.setValue("right_mode", self._right_mode)
            self._apply_right_layout()
            self.mode_changed.emit(self._right_mode)

    def _apply_right_layout(self) -> None:
        self._sync_button_checks()
        if self._right_mode == "popout":
            self._matrix_card.setMaximumHeight(16777215)
            self._matrix_stack.setVisible(True)
            self._msg_context_banner.setVisible(self._has_banner_content)
            self._ensure_popout_window()
            return

        if self._inspector.parent() is not self._right_splitter:
            self._right_splitter.addWidget(self._inspector)

        if self._right_mode == "matrix_only":
            self._right_splitter.setOrientation(Qt.Vertical)
            self._matrix_card.setMaximumHeight(16777215)
            self._matrix_stack.setVisible(True)
            self._msg_context_banner.setVisible(self._has_banner_content)
            self._inspector.stack.setVisible(False)
            strip_h = scale_px(38)
            self._inspector.setMaximumHeight(strip_h)
            mat_h = max(scale_px(260), self._right_splitter.height() - strip_h, scale_px(520))
            self._right_splitter.setSizes([mat_h, strip_h])
        elif self._right_mode == "inspector_only":
            self._right_splitter.setOrientation(Qt.Vertical)
            self._matrix_stack.setVisible(False)
            self._msg_context_banner.setVisible(False)
            strip_h = scale_px(40)
            self._matrix_card.setMaximumHeight(strip_h)
            self._inspector.setMaximumHeight(16777215)
            self._inspector.stack.setVisible(True)
            insp_h = max(scale_px(260), self._right_splitter.height() - strip_h, scale_px(520))
            self._right_splitter.setSizes([strip_h, insp_h])
        elif self._right_mode == "split_h":
            self._right_splitter.setOrientation(Qt.Horizontal)
            self._matrix_card.setMaximumHeight(16777215)
            self._matrix_stack.setVisible(True)
            self._msg_context_banner.setVisible(self._has_banner_content)
            self._inspector.setMaximumHeight(16777215)
            self._inspector.stack.setVisible(True)
            self._right_splitter.setSizes(self._restore_right_sizes("split_h"))
        else:
            self._right_splitter.setOrientation(Qt.Vertical)
            self._matrix_card.setMaximumHeight(16777215)
            self._matrix_stack.setVisible(True)
            self._msg_context_banner.setVisible(self._has_banner_content)
            self._inspector.setMaximumHeight(16777215)
            self._inspector.stack.setVisible(True)
            self._right_splitter.setSizes(self._restore_right_sizes("split_v"))

    def _sync_button_checks(self) -> None:
        pairs = [
            (self.btn_layout_v, self._right_mode == "split_v"),
            (self.btn_layout_h, self._right_mode == "split_h"),
            (self.btn_layout_matrix_max, self._right_mode == "matrix_only"),
        ]
        if self.btn_layout_insp_max is not None:
            pairs.append((self.btn_layout_insp_max, self._right_mode == "inspector_only"))
        if self.btn_layout_popout is not None:
            pairs.append((self.btn_layout_popout, self._right_mode == "popout"))
        for btn, checked in pairs:
            btn.blockSignals(True)
            btn.setChecked(checked)
            btn.blockSignals(False)

    def _save_right_splitter_sizes(self) -> None:
        if self._auto_compacted or self._right_mode not in ("split_v", "split_h"):
            return
        sizes = self._right_splitter.sizes()
        if len(sizes) >= 2 and all(s >= scale_px(80) for s in sizes[:2]):
            key = "sizes_h" if self._right_mode == "split_h" else "sizes_v"
            self._settings.setValue(key, f"{sizes[0]},{sizes[1]}")

    def _restore_right_sizes(self, mode: str) -> List[int]:
        key = "sizes_h" if mode == "split_h" else "sizes_v"
        raw = str(self._settings.value(key, "") or "")
        parts = [int(x) for x in raw.split(",") if x.strip().isdigit()]
        min_p = scale_px(180) if mode == "split_h" else scale_px(120)
        if len(parts) == 2 and parts[0] >= min_p and parts[1] >= min_p:
            return parts
        if mode == "split_h":
            return [scale_px(440), scale_px(560)]
        return [scale_px(360), scale_px(340)]

    def _on_right_splitter_moved(self, _pos: int, _index: int) -> None:
        self._auto_compacted = False
        self._save_right_splitter_sizes()

    def _on_main_splitter_moved(self, _pos: int, _index: int) -> None:
        if not self._left_collapsed:
            sizes = self._main_splitter.sizes()
            if len(sizes) >= 2 and sizes[0] >= scale_px(220):
                self._settings.setValue("left_width", sizes[0])

    def _on_left_splitter_moved(self, _pos: int, _index: int) -> None:
        if not self._specs_collapsed:
            sizes = self._left_splitter.sizes()
            if len(sizes) >= 2 and sizes[0] >= scale_px(120):
                self._settings.setValue("left_splitter_sizes", f"{sizes[0]},{sizes[1]}")

    def shutdown(self) -> None:
        """Persist splitter states and close any popped-out inspector cleanly."""
        self._shutting_down = True
        self._save_right_splitter_sizes()
        if self._popout_win is not None:
            pop = self._popout_win
            self._popout_win = None
            if self._inspector.parent() is pop.body_container:
                self._right_splitter.addWidget(self._inspector)
            pop.close_without_redock()
