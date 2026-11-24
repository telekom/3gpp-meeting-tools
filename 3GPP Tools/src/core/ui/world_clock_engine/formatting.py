"""Status bar pill text and rich multi-line tooltip formatters for WorldClockService."""

from __future__ import annotations

from typing import List

from core.ui.world_clock_engine.models import (
    WorldClockSnapshot,
    _format_wg_summary,
    _short_city_name,
)
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset


class WorldClockFormattingMixin:
    """Provides compact status bar pill text and rich tooltip formatters."""

    @staticmethod
    def format_status_bar_text(snapshot: WorldClockSnapshot, compact: bool = False) -> str:
        """Formats the compact status-bar clock pill text."""
        parts = [f"🕒 {snapshot.laptop_dt:%H:%M} Local"]
        if snapshot.show_utc_in_bar:
            parts.append(f"🌐 {snapshot.utc_dt:%H:%M} UTC")

        venue_chunks: List[str] = []
        for v in snapshot.primary_venues:
            short_loc = _short_city_name(v.location)
            if v.venue_dt is not None:
                day_tag = (
                    f" ({v.day_diff_vs_laptop:+d}d)"
                    if v.day_diff_vs_laptop != 0
                    else ""
                )
                chunk = f"{v.venue_dt:%H:%M} {short_loc}{day_tag}"
            else:
                chunk = f"--:-- {short_loc}"
            venue_chunks.append(chunk)

        if venue_chunks:
            prefix = "⏭️" if all(v.is_upcoming for v in snapshot.primary_venues) else "📍"
            if compact and len(venue_chunks) > 2:
                extra = len(venue_chunks) - 2
                shown_venues = " | ".join(venue_chunks[:2]) + f" +{extra}"
            else:
                shown_venues = " | ".join(venue_chunks)

            wg_suffix = ""
            if not compact and snapshot.wg_filter:
                wg_txt = "/".join(snapshot.wg_filter[:2])
                if len(snapshot.wg_filter) > 2:
                    wg_txt += "+"
                wg_suffix = f" ({wg_txt})"
            parts.append(f"{prefix} {shown_venues}{wg_suffix}")

        return " · ".join(parts)

    @staticmethod
    def format_status_bar_tooltip(snapshot: WorldClockSnapshot) -> str:
        """Formats a rich multi-line tooltip detailing all clocks, WGs, and sources."""
        lines = [
            f"💻 Laptop Local (IT): {snapshot.laptop_dt:%a %d %b %Y, %H:%M:%S} ({format_utc_offset(snapshot.laptop_offset_minutes)})",
            f"🌐 Official 3GPP UTC: {snapshot.utc_dt:%a %d %b %Y, %H:%M:%S} UTC",
        ]

        source_labels = {
            "manual_override": "Manual Override",
            "web_api": "Web Verified",
            "meeting_folder": "Meeting Folder",
            "builtin": "Built-in TZ",
            "unresolved": "Unresolved",
        }

        if snapshot.primary_venues:
            header = (
                "\n⏭️ Next Upcoming 3GPP Venue(s):"
                if all(v.is_upcoming for v in snapshot.primary_venues)
                else "\n📍 Active 3GPP Meeting Venue(s):"
            )
            lines.append(header)
            for v in snapshot.primary_venues:
                wg_txt = _format_wg_summary(v.wgs, max_items=14)
                src_txt = source_labels.get(v.source, v.source)
                if v.venue_dt is not None:
                    day_note = f" [{v.day_diff_vs_laptop:+d}d]" if v.day_diff_vs_laptop != 0 else ""
                    lines.append(
                        f"  • {v.location}: {v.venue_dt:%a %d %b, %H:%M:%S}{day_note} "
                        f"({format_utc_offset(v.tz_offset_minutes or 0)}, {src_txt})"
                    )
                else:
                    lines.append(f"  • {v.location}: Timezone unresolved")
                if wg_txt:
                    lines.append(f"    WGs: {wg_txt}")

        if snapshot.other_venues:
            lines.append("\n🌐 Other Active 3GPP Venues (Filtered):")
            for v in snapshot.other_venues:
                wg_txt = _format_wg_summary(v.wgs, max_items=10)
                if v.venue_dt is not None:
                    lines.append(
                        f"  • {v.location}: {v.venue_dt:%H:%M} ({format_utc_offset(v.tz_offset_minutes or 0)}) — {wg_txt}"
                    )

        if snapshot.wg_filter:
            lines.append(f"\n🎯 WG Filter Active: {', '.join(snapshot.wg_filter)}")

        lines.append("\nClick to configure WG filter, web sync, or manual overrides")
        return "\n".join(lines)
