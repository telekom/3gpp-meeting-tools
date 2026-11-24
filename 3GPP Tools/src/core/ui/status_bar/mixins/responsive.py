# --- File: src/core/ui/status_bar/mixins/responsive.py ---
"""Content-driven responsive width calculation and minimalist icon-mode switching."""

from typing import Optional

from PyQt5.QtWidgets import QWidget

from ..styles import _MINIMAL_HYSTERESIS_PX


class StatusBarResponsiveMixin:
    """Mixin handling minimalist mode sizing, hysteresis, and full-mode pill labels."""

    @property
    def is_minimal_mode(self) -> bool:
        """Whether the status bar is currently in narrow-window minimalist icon mode."""
        return self._minimal_mode

    def _full_update_text(self) -> str:
        visible, count, _, _ = self._update_state
        if not visible:
            return ""
        return f"Update ({count})" if count > 0 else "Update Available"

    def _full_ollama_text(self) -> str:
        if self._ollama_state is None:
            return f"🦙 {self._ollama_initial_model}" if self._ollama_initial_model else "🦙 Ollama"
        is_online, model_display, _, _ = self._ollama_state
        return f"🦙 {model_display}" if is_online else "🦙 Offline"

    def _full_mcp_text(self) -> str:
        active, total_calls = self._mcp_state
        if active:
            return f"🤖 MCP: 🟢 Active ({total_calls})"
        return f"🤖 MCP: {total_calls} calls" if total_calls > 0 else "🤖 MCP: Idle"

    def _full_proxy_text(self) -> str:
        return "🛡️ Proxy: ON" if self._proxy_active else "🌐 Direct"

    def _full_network_text(self) -> str:
        ssid, is_3gpp, server_reachable = self._network_state
        if is_3gpp and server_reachable:
            return f"🟢 {ssid} (Local Server Active)"
        if is_3gpp and not server_reachable:
            return f"🟡 {ssid} (No Local Server)"
        display_name = ssid if ssid else ("Network" if self.wifi_monitor is None else "Offline")
        return "📶 Network" if display_name == "Network" else f"🌐 {display_name}"

    def _full_drawer_text(self) -> str:
        parts = ["Console && Queue"]
        if self._drawer_queue_count > 0:
            parts.append(f"{self._drawer_queue_count} queued")
        if self._drawer_unread_errors > 0:
            noun = "err" if self._drawer_unread_errors == 1 else "errs"
            parts.append(f"{self._drawer_unread_errors} {noun}")
        return " · ".join(parts)

    def _estimate_full_pill_width(
        self, widget: QWidget, full_text: str, is_label: bool = False, has_icon: bool = False
    ) -> int:
        """Return exact/estimated pixel width of a pill when rendered in full mode."""
        if not self._minimal_mode:
            w = widget.sizeHint().width()
            self._full_pill_widths[widget] = (full_text, w)
            return w
        cached = self._full_pill_widths.get(widget)
        if cached is not None and cached[0] == full_text:
            return cached[1]
        display_text = full_text.replace("&&", "&")
        fm = widget.fontMetrics()
        text_w = fm.horizontalAdvance(display_text) if display_text else 0
        if is_label:
            return text_w + 16
        icon_w = 18 if has_icon else 0
        return text_w + icon_w + 18

    def compute_full_required_width(self) -> int:
        """
        Compute the total width required to display the left status message
        plus all active status bar pills in full (non-minimalist) mode.
        """
        entries = []
        if self._update_state[0]:
            entries.append((self.update_indicator, self._full_update_text(), False, True))
        entries.extend([
            (self.ollama_indicator, self._full_ollama_text(), False, False),
            (self.mcp_indicator, self._full_mcp_text(), False, False),
            (self.proxy_indicator, self._full_proxy_text(), False, False),
            (self.network_indicator, self._full_network_text(), True, False),
            (self.system_btn, "System", False, True),
            (self.assistant_dock_btn, "Assistant", False, True),
            (self.search_dock_btn, "Quick Search", False, True),
            (self.drawer_toggle_btn, self._full_drawer_text(), False, True),
            (self.clock_indicator, self._last_full_clock_text, False, False),
        ])

        pills_total = sum(
            self._estimate_full_pill_width(w, txt, is_label=is_lbl, has_icon=has_ico)
            for w, txt, is_lbl, has_ico in entries
        )
        spacing_total = len(entries) * 6

        msg = self.currentMessage() or "🟢 System Idle."
        msg_w = self.fontMetrics().horizontalAdvance(msg)
        left_reserve = max(125, msg_w + 28)

        return left_reserve + pills_total + spacing_total + 24

    def _on_status_message_changed(self, _message: str) -> None:
        self._recheck_responsive_fit()

    def _recheck_responsive_fit(self) -> None:
        """Re-evaluate minimalist mode after pill content or left message changes."""
        if not self._responsive_armed or self._in_responsive_check:
            return
        self.apply_responsive_width(self._last_available_width)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        w = event.size().width()
        if w > 0:
            self._responsive_armed = True
            self.apply_responsive_width(w)

    def apply_responsive_width(self, available_width: Optional[int] = None) -> None:
        """Dynamically toggle minimalist mode based on actual full-mode pill width + hysteresis."""
        if self._in_responsive_check:
            return
        if available_width is not None:
            self._responsive_armed = True
            self._last_available_width = int(available_width)
            avail = int(available_width)
        elif self.width() > 200 and self.isVisible():
            avail = self.width()
            self._last_available_width = avail
        elif self._last_available_width is not None:
            avail = self._last_available_width
        else:
            win = self.window()
            avail = win.width() if win is not None else self.width()

        if avail <= 0:
            return

        self._in_responsive_check = True
        try:
            req = self.compute_full_required_width()
            if not self._minimal_mode and avail < req:
                self.set_minimal_mode(True)
            elif self._minimal_mode and avail >= req + _MINIMAL_HYSTERESIS_PX:
                self.set_minimal_mode(False)
        finally:
            self._in_responsive_check = False

    def set_minimal_mode(self, minimal: bool) -> None:
        """Switch between full pill labels and compact icon-only minimalist mode."""
        want = bool(minimal)
        if want == self._minimal_mode:
            return
        self._minimal_mode = want
        self._refresh_all_indicators()

    def _refresh_all_indicators(self) -> None:
        """Re-render every status bar pill for the current minimalist mode state."""
        self._render_update_indicator()
        self._render_ollama_indicator()
        self._render_mcp_indicator()
        self._render_proxy_indicator()
        self._render_network_indicator()
        self._render_system_btn()
        self.update_right_panel_badges(self._right_panel_open, self._right_panel_mode)
        self.update_drawer_badge(
            self._drawer_open,
            self._drawer_queue_count,
            self._drawer_unread_errors,
        )
        self.update_clock_indicator()
