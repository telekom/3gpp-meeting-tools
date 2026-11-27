"""Side-panel splitter management and anchor routing mixin for PacketStackWidget."""

from __future__ import annotations

from urllib.parse import unquote

from PyQt5.QtCore import Qt, QUrl

from .packet_ie_resolver import resolve_packet_field


class PacketStackSidePanelMixin:
    """Manages `body_splitter`, inline `side_panel`, dock orientation, and packet link clicks."""

    def _version_hint_for_spec(self, spec_number: str) -> str:
        clean = (spec_number or "").replace("TS", "").strip()
        if clean == "38.415":
            return self._config.spec_version_38415
        if clean == "29.281":
            return self._config.spec_version_29281
        if clean == "38.413":
            return self._config.spec_version_38413
        if clean == "24.501":
            return self._config.spec_version_24501
        if clean in ("24.526", "24.588", "24.555", "24.514", "24.578"):
            return self._config.spec_version_policy
        return ""

    def _apply_initial_splitter_sizes(self) -> None:
        raw = self._settings.value("packet_stack/side_panel_sizes")
        if isinstance(raw, list) and len(raw) == 2:
            try:
                sizes = [int(raw[0]), int(raw[1])]
                if sizes[0] > 80 and sizes[1] > 80:
                    self.body_splitter.setSizes(sizes)
                    return
            except (TypeError, ValueError):
                pass
        span = self.body_splitter.width() if self.body_splitter.orientation() == Qt.Horizontal else self.body_splitter.height()
        total = max(600, span)
        self.body_splitter.setSizes([int(total * 0.6), int(total * 0.4)])

    def _on_body_splitter_moved(self, _pos: int, _index: int) -> None:
        if not self.side_panel.isVisible():
            return
        sizes = self.body_splitter.sizes()
        if len(sizes) == 2 and sizes[0] > 0 and sizes[1] > 0:
            self._settings.setValue("packet_stack/side_panel_sizes", sizes)

    def show_field_in_side_panel(
        self,
        spec_number: str,
        clause: str,
        field_name: str,
        *,
        push_history: bool = False,
    ) -> None:
        """Resolve and display a packet field/IE inline in the collapsible Side Panel."""
        res = resolve_packet_field(
            self.db,
            self._composed,
            spec_number,
            clause,
            field_name,
            version_hint=self._version_hint_for_spec(spec_number),
        )
        self.side_panel.show_result(res, push_history=push_history)
        if self.side_panel.isHidden():
            self.side_panel.setVisible(True)
            self._apply_initial_splitter_sizes()

    def hide_side_panel(self) -> None:
        self.side_panel.setVisible(False)

    def toggle_side_panel_dock(self) -> None:
        is_horiz = self.body_splitter.orientation() == Qt.Horizontal
        new_orient = Qt.Vertical if is_horiz else Qt.Horizontal
        self.body_splitter.setOrientation(new_orient)
        self.side_panel.set_docked_bottom(new_orient == Qt.Vertical)
        self._settings.setValue("packet_stack/side_panel_orientation", "v" if new_orient == Qt.Vertical else "h")
        if self.side_panel.isVisible():
            span = self.body_splitter.width() if new_orient == Qt.Horizontal else self.body_splitter.height()
            total = max(600, span)
            self.body_splitter.setSizes([int(total * 0.58), int(total * 0.42)])

    def _on_anchor_clicked(self, url: QUrl) -> None:
        url_str = url.toString()
        if url_str.startswith("action:open_spec:"):
            parts = [unquote(p) for p in url_str.replace("action:open_spec:", "").split(":")]
            spec_num = parts[0] if len(parts) > 0 else ""
            clause = parts[2] if len(parts) > 2 else ""
            self.show_field_in_side_panel(spec_num, clause, "")
            return

        if url_str.startswith("action:packet_field:"):
            parts = [unquote(p) for p in url_str.replace("action:packet_field:", "").split(":", 2)]
            spec_num = parts[0] if len(parts) > 0 else ""
            clause = parts[1] if len(parts) > 1 else ""
            field_name = parts[2] if len(parts) > 2 else ""
            self.show_field_in_side_panel(spec_num, clause, field_name)
