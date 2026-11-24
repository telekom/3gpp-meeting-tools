"""Live tick refresh, manual override persistence, web/IP sync workers, and shutdown mixin."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from core.ui.world_clock_service import VenueClockState, WebTimezoneSyncWorker
from modules.meetings.core.session_plan.session_plan_timezone import (
    SessionPlanTimezone,
    format_utc_offset,
)


class WorldClockSyncActionsMixin:
    """Handles snapshot ticks, manual timezone overrides, async Geo-IP/city sync, and teardown."""

    def _on_tick(self) -> None:
        self._refresh_snapshot(rebuild_cards=False)

    def _refresh_snapshot(self, rebuild_cards: bool = False) -> None:
        snapshot = self.service.get_snapshot()
        self._update_reference_clocks(snapshot)

        focused_state: Optional[VenueClockState] = None
        if self.focus_meeting is not None:
            focused_state = self.service.resolve_single_meeting(
                self.focus_meeting, meeting_dir=self.focus_meeting_dir, ref_utc=snapshot.utc_dt
            )

        sig = (
            focused_state.location.lower() if focused_state else "",
            tuple((v.location.lower(), v.is_upcoming) for v in snapshot.primary_venues),
            tuple((v.location.lower(), v.is_upcoming) for v in snapshot.other_venues),
        )

        if rebuild_cards or sig != self._last_layout_signature:
            self._rebuild_venue_cards(snapshot, focused_state)
            self._last_layout_signature = sig
        else:
            self._update_existing_cards(snapshot, focused_state)

    def _apply_manual_offset(
        self,
        location: str,
        offset_min: int,
        meetings: list,
        explicit_dir: Optional[Path],
    ) -> None:
        label = f"{format_utc_offset(offset_min)} (Manual Override)"
        if explicit_dir is not None and self.focus_meeting is not None:
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=explicit_dir,
                tz_offset_minutes=offset_min,
                tz_label=label,
                source="manual_override",
                location=location,
                wg=str(self.focus_meeting.get("wg_name", "")),
                meeting_number=str(self.focus_meeting.get("meeting_number", "")),
            )
        self.service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="manual_override",
            matching_meetings=meetings,
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_offset_override(self, location: str, offset_min: int, meetings: list, explicit_dir) -> None:
        self._apply_manual_offset(location, offset_min, meetings, explicit_dir)

    def _on_nudge_offset(self, location: str, delta_min: int, meetings: list, explicit_dir) -> None:
        cur_off, _lbl, _iana, _src = self.service.resolve_venue_timezone(
            location, meetings=meetings, explicit_meeting_dir=explicit_dir
        )
        base = cur_off if cur_off is not None else 0
        new_off = max(-12 * 60, min(14 * 60, int(base) + int(delta_min)))
        self._apply_manual_offset(location, new_off, meetings, explicit_dir)

    def _on_reset_venue(self, location: str, meetings: list, explicit_dir) -> None:
        if explicit_dir is not None:
            SessionPlanTimezone.clear_meeting_folder_timezone(Path(explicit_dir))
        self.service.settings.clear_venue_timezone_override(
            location=location, matching_meetings=meetings
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_web_sync_venue(self, location: str, meetings: list, explicit_dir) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == location.lower():
                card.set_syncing(True)

        worker = WebTimezoneSyncWorker(mode="city", location=location, meetings=meetings, parent=self)
        worker.city_resolved.connect(
            lambda loc, off, lbl, iana, mtgs, ed=explicit_dir: self._on_city_sync_done(
                loc, off, lbl, iana, mtgs, ed
            )
        )
        worker.sync_failed.connect(self._on_city_sync_failed)
        worker.finished.connect(lambda w=worker: self._cleanup_worker(w))
        self._active_workers.append(worker)
        worker.start()

    def _on_sync_all_venues(self) -> None:
        seen = set()
        for loc_key, _isf, card in self._venue_cards:
            if loc_key in seen:
                continue
            seen.add(loc_key)
            self._on_web_sync_venue(card.venue.location, card.venue.meetings, card.explicit_meeting_dir)

    def _on_city_sync_done(
        self,
        location: str,
        offset_min: int,
        label: str,
        iana_tz: str,
        meetings: list,
        explicit_dir: Optional[Path],
    ) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == location.lower():
                card.set_syncing(False)

        # Explicit user Web Sync replaces any prior manual lock so fresh web verification takes effect
        self.service.settings.clear_venue_timezone_override(
            location=location, matching_meetings=meetings
        )
        if explicit_dir is not None:
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=Path(explicit_dir),
                tz_offset_minutes=offset_min,
                tz_label=label,
                source="web_api",
                location=location,
                wg=str((self.focus_meeting or {}).get("wg_name", "")),
                meeting_number=str((self.focus_meeting or {}).get("meeting_number", "")),
                iana_tz=iana_tz,
            )

        self.service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="web_api",
            iana_tz=iana_tz,
            matching_meetings=meetings,
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_city_sync_failed(self, target: str, err: str) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == (target or "").lower():
                card.set_syncing(False)
        self.ip_status_lbl.setText(f"⚠️ Web sync failed for {target}: {err}")

    def _on_detect_ip_clicked(self) -> None:
        self.btn_detect_ip.setEnabled(False)
        self.btn_detect_ip.setText("⏳ Detecting...")
        self.btn_apply_ip.setVisible(False)
        self.ip_status_lbl.setText("Querying public Geo-IP timezone server...")

        worker = WebTimezoneSyncWorker(mode="ip", parent=self)
        worker.ip_resolved.connect(self._on_ip_detected)
        worker.sync_failed.connect(self._on_ip_failed)
        worker.finished.connect(lambda w=worker: self._cleanup_worker(w))
        self._active_workers.append(worker)
        worker.start()

    def _on_ip_detected(self, city: str, offset_min: int, label: str, iana_tz: str) -> None:
        self.btn_detect_ip.setEnabled(True)
        self.btn_detect_ip.setText("📡 Detect My Physical Timezone (via Public IP)")
        self._detected_ip_info = (city, offset_min, label, iana_tz)
        self.ip_status_lbl.setText(f"Detected: {city} — {label}")
        self.btn_apply_ip.setVisible(bool(self._venue_cards))

    def _on_ip_failed(self, _target: str, err: str) -> None:
        self.btn_detect_ip.setEnabled(True)
        self.btn_detect_ip.setText("📡 Detect My Physical Timezone (via Public IP)")
        self.ip_status_lbl.setText(f"⚠️ Could not detect IP timezone: {err}")

    def _on_apply_ip_to_primary(self) -> None:
        if not self._detected_ip_info or not self._venue_cards:
            return
        _city, offset_min, _label, _iana = self._detected_ip_info
        _loc_key, _isf, card = self._venue_cards[0]
        self._apply_manual_offset(
            card.venue.location, offset_min, card.venue.meetings, card.explicit_meeting_dir
        )

    def _cleanup_worker(self, worker: WebTimezoneSyncWorker) -> None:
        if worker in self._active_workers:
            self._active_workers.remove(worker)

    def closeEvent(self, event) -> None:
        if hasattr(self, "_tick_timer") and self._tick_timer.isActive():
            self._tick_timer.stop()
        for w in list(self._active_workers):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        self._active_workers.clear()
        super().closeEvent(event)

    def done(self, result: int) -> None:
        if hasattr(self, "_tick_timer") and self._tick_timer.isActive():
            self._tick_timer.stop()
        for w in list(self._active_workers):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        self._active_workers.clear()
        super().done(result)
