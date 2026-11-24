# --- File: src/core/ui/status_bar/mixins/ai_and_clock.py ---
"""Ollama health monitoring, MCP cockpit polling, World Clock pill, and shutdown cleanup."""

import logging

from core.ai.ollama_client import OllamaMonitorThread, load_ollama_config
from core.ai.ollama_dialog import OllamaConfigDialog
from core.ui.theme import C
from core.ui.world_clock_service import WebTimezoneSyncWorker

from ..styles import _status_pill_qss

logger = logging.getLogger(__name__)


class StatusBarAiClockMixin:
    """Mixin managing Ollama monitor, MCP activity polling, World Clock, and clean shutdown."""

    def start_ollama_monitor(self):
        """Starts Ollama monitor independently."""
        try:
            if self.ollama_monitor is None or not self.ollama_monitor.isRunning():
                logger.info("🏁 [STARTUP:SERVICES] Starting OllamaMonitorThread...")
                self.ollama_monitor = OllamaMonitorThread(parent=None)
                self.ollama_monitor.status_updated.connect(self._update_ollama_indicator)
                self.ollama_monitor.start()
        except Exception as e:
            logger.error(f"❌ Failed to start OllamaMonitorThread: {e}")
            self._ollama_state = (False, "", 0, str(e))
            self._render_ollama_indicator()

    def _render_ollama_indicator(self) -> None:
        if self._ollama_state is None:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))
            return

        is_online, model_display, models_count, error = self._ollama_state
        if is_online:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setToolTip(
                f"Ollama Online ({model_display}, {models_count} models available)\nClick to configure"
            )
            self.ollama_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.SUCCESS_SOFT,
                    fg=C.SUCCESS,
                    bd=C.SUCCESS_BORDER,
                    hover_bg=C.SUCCESS_SOFT_ALT,
                    hover_bd=C.SUCCESS_BORDER_MINT,
                    compact=self._minimal_mode,
                )
            )
        else:
            self.ollama_indicator.setText("🦙" if self._minimal_mode else self._full_ollama_text())
            self.ollama_indicator.setToolTip(f"Ollama Unreachable: {error}\nClick to configure")
            self.ollama_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.DANGER_SOFT,
                    fg=C.DANGER_ALT,
                    bd=C.DANGER_BORDER,
                    hover_bg=C.DANGER_SOFT_HOVER,
                    hover_bd=C.DANGER_BORDER_ALT,
                    compact=self._minimal_mode,
                )
            )

    def _update_ollama_indicator(self, is_online: bool, selected_model: str, models: list, error: str):
        model_display = selected_model if selected_model else (models[0] if models else "Connected")
        self._ollama_state = (bool(is_online), str(model_display), len(models), str(error or ""))
        self._render_ollama_indicator()
        self._recheck_responsive_fit()

    def open_ollama_settings(self):
        """Opens Ollama connection, model preferences, and service management dialog."""
        dialog = OllamaConfigDialog(self.window())
        dialog.exec_()

        if self.ollama_monitor and self.ollama_monitor.isRunning():
            cfg = load_ollama_config()
            self.ollama_monitor.client.reconfigure(cfg.get("host"), cfg.get("proxy_mode"))
            self.ollama_monitor.trigger_check()

    def open_mcp_monitor(self):
        """Opens the MCP Server & External AI Agent Cockpit dialog."""
        try:
            from core.ui.mcp_monitor_dialog import McpMonitorDialog
            if self._mcp_dialog is None:
                self._mcp_dialog = McpMonitorDialog(self.window())
            self._mcp_dialog.show()
            self._mcp_dialog.raise_()
            self._mcp_dialog.activateWindow()
            self._mcp_dialog.refresh_all()
        except Exception as exc:
            logger.exception("Failed to open MCP Monitor dialog: %s", exc)

    def _render_mcp_indicator(self) -> None:
        active, total_calls = self._mcp_state
        if active:
            text = "🤖 🟢" if self._minimal_mode else f"🤖 MCP: 🟢 Active ({total_calls})"
            self.mcp_indicator.setText(text)
            self.mcp_indicator.setStyleSheet(
                _status_pill_qss(
                    bg=C.SUCCESS_SOFT,
                    fg=C.SUCCESS,
                    bd=C.SUCCESS_BORDER,
                    hover_bg=C.SUCCESS_SOFT_ALT,
                    hover_bd=C.SUCCESS_BORDER_MINT,
                    compact=self._minimal_mode,
                )
            )
        else:
            if self._minimal_mode:
                text = "🤖"
            else:
                text = f"🤖 MCP: {total_calls} calls" if total_calls > 0 else "🤖 MCP: Idle"
            self.mcp_indicator.setText(text)
            self.mcp_indicator.setStyleSheet(_status_pill_qss(compact=self._minimal_mode))

    def update_mcp_indicator(self):
        """Polls active MCP sessions and recent activity stats to update the status bar pill."""
        try:
            from core.ai.mcp_activity import get_active_mcp_sessions, get_mcp_activity_stats
            sessions = get_active_mcp_sessions()
            stats = get_mcp_activity_stats()
            total_calls = stats.get("total_calls", 0)

            if sessions:
                clients = ", ".join(s.get("client", "agent") for s in sessions)
                pids = ", ".join(str(s.get("pid", "?")) for s in sessions)
                tooltip = (
                    f"MCP Server: Active\n"
                    f"Connected Client: {clients} (PID: {pids})\n"
                    f"Total Calls: {total_calls}\n"
                    f"Success Rate: {stats.get('success_rate', 100)}%\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self._mcp_state = (True, total_calls)
                self.mcp_indicator.setToolTip(tooltip)
                self._render_mcp_indicator()
            else:
                tooltip = (
                    f"MCP Server: Idle (No active agent process connected)\n"
                    f"Total Calls Logged: {total_calls}\n\n"
                    f"Click to open MCP & AI Agent Cockpit"
                )
                self._mcp_state = (False, total_calls)
                self.mcp_indicator.setToolTip(tooltip)
                self._render_mcp_indicator()
            self._recheck_responsive_fit()
        except Exception:
            pass

    def update_clock_indicator(self) -> None:
        """Refreshes the status-bar World Clock pill text and rich tooltip."""
        try:
            snapshot = self._world_clock_service.get_snapshot()
            self._last_full_clock_text = self._world_clock_service.format_status_bar_text(
                snapshot, compact=False
            )
            self.clock_indicator.setText(
                self._world_clock_service.format_status_bar_text(
                    snapshot, compact=self._minimal_mode
                )
            )
            self.clock_indicator.setToolTip(self._world_clock_service.format_status_bar_tooltip(snapshot))
            self.clock_indicator.setStyleSheet(
                _status_pill_qss(bg=C.BG_SURFACE, fg=C.TEXT_BODY, compact=self._minimal_mode)
            )
            self._recheck_responsive_fit()
        except Exception as e:
            logger.debug(f"[StatusBarClock] Failed to update clock indicator: {e}")

    def refresh_world_clock(self) -> None:
        """Invalidates meeting DB cache and immediately refreshes the clock pill."""
        self._world_clock_service.invalidate_cache()
        self.update_clock_indicator()

    def open_world_clock(self) -> None:
        """Opens the interactive World Clock & Active 3GPP Venue Timezone dialog."""
        from core.ui.world_clock_dialog import WorldClockDialog

        dlg = WorldClockDialog(self._world_clock_service, parent=self.window())
        dlg.settings_changed.connect(self.refresh_world_clock)
        dlg.exec_()
        self.refresh_world_clock()

    def _schedule_active_venue_web_sync(self) -> None:
        """Background-verifies active meeting venue timezones not yet web-verified or manually overridden."""
        if self._monitors_stopped:
            return
        try:
            snapshot = self._world_clock_service.get_snapshot()
            for v in snapshot.primary_venues + snapshot.other_venues:
                loc_key = v.location.strip().lower()
                if (
                    v.source in ("builtin", "unresolved")
                    and loc_key
                    and loc_key not in self._clock_auto_synced
                ):
                    self._clock_auto_synced.add(loc_key)
                    worker = WebTimezoneSyncWorker(
                        mode="city",
                        location=v.location,
                        meetings=v.meetings,
                        parent=self,
                    )
                    worker.city_resolved.connect(self._on_auto_venue_web_resolved)
                    worker.finished.connect(
                        lambda w=worker: self._clock_sync_workers.remove(w)
                        if w in self._clock_sync_workers
                        else None
                    )
                    self._clock_sync_workers.append(worker)
                    worker.start()
        except Exception as e:
            logger.debug(f"[StatusBarClock] Auto web sync scheduling failed: {e}")

    def _on_auto_venue_web_resolved(
        self, location: str, offset_min: int, label: str, iana_tz: str, meetings: list
    ) -> None:
        self._world_clock_service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="web_api",
            iana_tz=iana_tz,
            matching_meetings=meetings,
        )
        self.refresh_world_clock()

    def stop_monitors(self):
        """Cleanly stops background thread monitors upon window shutdown."""
        self._monitors_stopped = True
        if hasattr(self, "clock_timer") and self.clock_timer.isActive():
            self.clock_timer.stop()
        if hasattr(self, "mcp_timer") and self.mcp_timer.isActive():
            self.mcp_timer.stop()
        for w in list(getattr(self, "_clock_sync_workers", [])):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        if hasattr(self, "_clock_sync_workers"):
            self._clock_sync_workers.clear()
        if self.wifi_monitor and self.wifi_monitor.isRunning():
            self.wifi_monitor.requestInterruption()
            self.wifi_monitor.wait(300)
        if self.ollama_monitor and self.ollama_monitor.isRunning():
            self.ollama_monitor.requestInterruption()
            self.ollama_monitor.wait(300)
