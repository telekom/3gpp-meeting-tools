# --- File: src/modules/meetings/core/session_plan_exporter.py ---
"""
Multi-format exporter for 3GPP SA2 Meeting Session Plans.
1. Mobile-First Standalone HTML: Zero-dependency, offline-ready, responsive day/quarter cards.
2. iCalendar (.ics): RFC 5545 UTC-anchored events for Outlook / Apple / Google Calendar.
3. Colleague Email Summary: Pre-formatted executive HTML/text for Outlook paste.
"""
import html
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan_diff import SessionPlanDiffResult
from modules.meetings.core.session_plan_models import (
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
)
from modules.meetings.core.session_plan_timezone import SessionPlanTimezone

logger = logging.getLogger(__name__)


class SessionPlanExporter:
    """Generates mobile HTML, .ics calendar, and email summaries."""

    @classmethod
    def export_mobile_html(
        cls,
        snapshot: SessionPlanSnapshot,
        output_path: Path,
        meeting_tz_offset_min: int = 120,
        tz_name: str = "Venue Time",
        meeting_dates_str: str = "",
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
        tracked_tokens: Optional[Set[str]] = None,
    ) -> Path:
        """
        Exports a self-contained, responsive, 100% offline mobile HTML view.
        No external CDN dependencies (fonts and CSS are purely embedded).
        """
        annotations = annotations or {}
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        laptop_offset_min = SessionPlanTimezone.get_laptop_tz_offset_minutes()
        is_tz_different = (laptop_offset_min != meeting_tz_offset_min)

        # Group slots by Day and Quarter
        slots_by_day_q: Dict[str, Dict[str, List[SessionSlot]]] = {}
        for s in snapshot.slots:
            slots_by_day_q.setdefault(s.day_name, {}).setdefault(s.quarter, []).append(s)

        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        active_days = [d for d in day_names if d in slots_by_day_q]

        streams = snapshot.streams

        # Build HTML content
        html_parts = []
        html_parts.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>SA2 #{snapshot.meeting_number} Session Plan ({snapshot.tdoc_id})</title>
<style>
  :root {{
    --bg: #F8FAFC;
    --card: #FFFFFF;
    --text: #0F172A;
    --subtext: #475569;
    --primary: #2563EB;
    --primary-light: #EFF6FF;
    --stream1: #4F46E5;
    --stream2: #059669;
    --stream3: #D97706;
    --drafting: #DC2626;
    --drafting-bg: #FEF2F2;
    --tracked: #7C3AED;
    --tracked-bg: #F5F3FF;
    --border: #E2E8F0;
  }}
  * {{ box-sizing: border-box; -webkit-tap-highlight-color: transparent; }}
  body {{
    margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    font-size: 14px; line-height: 1.4;
  }}
  header {{
    position: sticky; top: 0; z-index: 100; background: #FFFFFF;
    border-bottom: 1px solid var(--border); padding: 10px 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
  }}
  .title-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }}
  .title {{ font-size: 16px; font-weight: 700; color: #1E293B; margin: 0; }}
  .badge-tdoc {{ font-size: 11px; background: #E0E7FF; color: #3730A3; padding: 2px 7px; border-radius: 999px; font-weight: 600; }}
  .tz-banner {{ font-size: 11px; color: var(--subtext); margin-bottom: 8px; }}
  .filter-bar {{ display: flex; gap: 8px; align-items: center; }}
  .search-box {{
    flex: 1; padding: 7px 10px; border: 1px solid var(--border); border-radius: 8px;
    font-size: 13px; outline: none; background: #F1F5F9;
  }}
  .search-box:focus {{ background: #FFFFFF; border-color: var(--primary); }}
  .btn-tracked {{
    padding: 6px 10px; border-radius: 8px; border: 1px solid #DDD6FE; background: #F5F3FF;
    color: #6D28D9; font-weight: 600; font-size: 12px; cursor: pointer; white-space: nowrap;
  }}
  .btn-tracked.active {{ background: #6D28D9; color: #FFFFFF; border-color: #6D28D9; }}
  .day-tabs {{
    display: flex; gap: 6px; overflow-x: auto; padding: 8px 14px; background: #FFFFFF;
    border-bottom: 1px solid var(--border); -webkit-overflow-scrolling: touch;
  }}
  .day-tab {{
    padding: 6px 14px; border-radius: 999px; font-weight: 600; font-size: 13px;
    background: #F1F5F9; color: var(--subtext); border: none; cursor: pointer; white-space: nowrap;
  }}
  .day-tab.active {{ background: var(--primary); color: #FFFFFF; }}
  .content {{ padding: 12px 14px 60px 14px; }}
  .quarter-card {{
    background: var(--card); border: 1px solid var(--border); border-radius: 12px;
    margin-bottom: 12px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03);
  }}
  .quarter-header {{
    background: #F8FAFC; padding: 8px 12px; border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
  }}
  .quarter-title {{ font-size: 13px; font-weight: 700; color: #1E293B; }}
  .quarter-time {{ font-size: 12px; color: var(--subtext); }}
  .slot-row {{
    padding: 10px 12px; border-bottom: 1px solid #F1F5F9;
  }}
  .slot-row:last-child {{ border-bottom: none; }}
  .slot-row.drafting {{ background: var(--drafting-bg); border-left: 4px solid var(--drafting); }}
  .slot-row.tracked {{ background: var(--tracked-bg); border-left: 4px solid var(--tracked); }}
  .stream-badge {{
    display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 6px;
    border-radius: 4px; text-transform: uppercase; margin-right: 6px;
  }}
  .stream-badge.str-1 {{ background: #EEF2FF; color: var(--stream1); }}
  .stream-badge.str-2 {{ background: #ECFDF5; color: var(--stream2); }}
  .stream-badge.str-3 {{ background: #FFFBEB; color: var(--stream3); }}
  .room-pill {{ font-size: 11px; color: var(--subtext); font-weight: 500; }}
  .sub-slot-badge {{
    font-size: 10px; font-weight: 600; padding: 1px 5px; border-radius: 3px;
    background: #E2E8F0; color: #334155; margin-left: 6px;
  }}
  .slot-content {{ margin-top: 5px; font-size: 13px; font-weight: 500; color: #1E293B; }}
  .ai-chip {{
    display: inline-block; font-size: 11px; font-weight: 600; padding: 1px 6px;
    border-radius: 4px; background: #E0E7FF; color: #3730A3; margin-right: 4px; margin-top: 2px;
  }}
  .count-chip {{
    font-size: 10px; font-weight: 700; padding: 1px 4px; border-radius: 3px;
    background: #F1F5F9; color: #475569; margin-left: 4px;
  }}
  .note-chip {{
    display: block; margin-top: 5px; padding: 4px 8px; border-radius: 6px;
    background: #FEF3C7; color: #92400E; font-size: 11px; font-weight: 500;
  }}
  .hidden {{ display: none !important; }}
</style>
</head>
<body>
<header>
  <div class="title-row">
    <h1 class="title">SA2 #{snapshot.meeting_number} Session Plan</h1>
    <span class="badge-tdoc">{snapshot.tdoc_id}</span>
  </div>
  <div class="tz-banner">
    📍 {tz_name}
    {" | 💻 Austrian Laptop Time: " + ("+0" if laptop_offset_min >= meeting_tz_offset_min else "") + str((laptop_offset_min - meeting_tz_offset_min)//60) + "h" if is_tz_different else ""}
  </div>
  <div class="filter-bar">
    <input type="text" id="searchInput" class="search-box" placeholder="🔍 Filter AI, topic, or drafting..." oninput="applyFilters()">
    <button id="trackedBtn" class="btn-tracked" onclick="toggleTracked()">⭐ My Topics</button>
  </div>
</header>

<div class="day-tabs">
""")

        # Add Day Tabs
        for idx, d_name in enumerate(active_days):
            active_cls = " active" if idx == 0 else ""
            html_parts.append(f'<button class="day-tab{active_cls}" onclick="switchDay(\'{d_name}\')">{d_name[:3]}</button>\n')

        html_parts.append('</div>\n<div class="content">\n')

        # Add Content per Day
        for idx, d_name in enumerate(active_days):
            day_display = "" if idx == 0 else " style=\"display:none;\""
            html_parts.append(f'<div id="day-{d_name}" class="day-section"{day_display}>\n')

            q_map = slots_by_day_q.get(d_name, {})
            # Standard quarter order
            q_order = ["Q0", "Q1", "Q2", "Lunch", "Q3", "Q4", "Q5"]
            for q in q_order:
                if q not in q_map:
                    continue
                slots_in_q = q_map[q]
                if not slots_in_q:
                    continue

                # Quarter Header
                sample_s = slots_in_q[0]
                q_time_str = SessionPlanTimezone.format_dual_time_range(
                    sample_s.start_time_local, sample_s.end_time_local, meeting_tz_offset_min
                )

                html_parts.append(f"""
  <div class="quarter-card">
    <div class="quarter-header">
      <span class="quarter-title">{q}</span>
      <span class="quarter-time">{q_time_str}</span>
    </div>
""")

                # Render slots
                for s in slots_in_q:
                    s_info = streams.get(s.stream_number)
                    r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"
                    convener = s_info.convener if s_info else ""

                    # Check if tracked
                    is_tracked = False
                    for it in s.items:
                        if it.ai_number.lower() in tracked_tokens or any(tk in it.topic_label.lower() for tk in tracked_tokens):
                            is_tracked = True
                            break

                    classes = ["slot-row"]
                    if s.is_drafting:
                        classes.append("drafting")
                    if is_tracked:
                        classes.append("tracked")

                    # User notes
                    note_obj = annotations.get(s.slot_key)
                    note_html = ""
                    if note_obj and note_obj.comment:
                        note_html = f'<div class="note-chip">📝 <b>Note:</b> {html.escape(note_obj.comment)}</div>'

                    # Sub-slot tag
                    sub_tag = ""
                    if s.sub_slot == 1:
                        sub_tag = f'<span class="sub-slot-badge">TU 1 ({s.start_time_local}–{s.end_time_local})</span>'
                    elif s.sub_slot == 2:
                        sub_tag = f'<span class="sub-slot-badge">TU 2 ({s.start_time_local}–{s.end_time_local})</span>'

                    # Item chips
                    chips = []
                    for it in s.items:
                        count_html = f'<span class="count-chip">[{it.tdoc_count_str}]</span>' if it.tdoc_count_str else ""
                        ai_label = f"AI {it.ai_number}" if it.ai_number else it.topic_label
                        chips.append(f'<span class="ai-chip">{html.escape(ai_label)}{count_html}</span>')

                    chips_html = "".join(chips)
                    raw_esc = html.escape(s.raw_text)

                    html_parts.append(f"""
    <div class="{' '.join(classes)}" data-text="{raw_esc.lower()}" data-tracked="{1 if is_tracked else 0}">
      <div>
        <span class="stream-badge str-{s.stream_number}">Stream {s.stream_number}</span>
        <span class="room-pill">📍 {html.escape(r_name)} {"· " + html.escape(convener) if convener else ""}</span>
        {sub_tag}
      </div>
      <div class="slot-content">
        {raw_esc}
      </div>
      {f'<div style="margin-top:3px;">{chips_html}</div>' if chips_html else ''}
      {note_html}
    </div>
""")

                html_parts.append("  </div>\n")  # Close quarter-card

            html_parts.append("</div>\n")  # Close day-section

        # Embedded Vanilla JavaScript
        html_parts.append("""
</div>
<script>
  let onlyTracked = false;

  function switchDay(dayName) {
    document.querySelectorAll('.day-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.day-section').forEach(s => s.style.display = 'none');
    event.target.classList.add('active');
    const target = document.getElementById('day-' + dayName);
    if (target) target.style.display = 'block';
  }

  function toggleTracked() {
    onlyTracked = !onlyTracked;
    const btn = document.getElementById('trackedBtn');
    btn.classList.toggle('active', onlyTracked);
    applyFilters();
  }

  function applyFilters() {
    const q = document.getElementById('searchInput').value.toLowerCase().trim();
    document.querySelectorAll('.slot-row').forEach(row => {
      const txt = row.getAttribute('data-text') || '';
      const isTrk = row.getAttribute('data-tracked') === '1';
      const matchesSearch = !q || txt.includes(q);
      const matchesTracked = !onlyTracked || isTrk;

      if (matchesSearch && matchesTracked) {
        row.classList.remove('hidden');
      } else {
        row.classList.add('hidden');
      }
    });

    // Hide empty quarter cards
    document.querySelectorAll('.quarter-card').forEach(card => {
      const visibleSlots = card.querySelectorAll('.slot-row:not(.hidden)');
      card.style.display = visibleSlots.length > 0 ? 'block' : 'none';
    });
  }
</script>
</body>
</html>
""")

        output_path.write_text("".join(html_parts), encoding="utf-8")
        logger.info(f"[SessionPlanExporter] Generated mobile HTML: {output_path} ({output_path.stat().st_size} bytes)")
        return output_path

    @classmethod
    def export_ics(
        cls,
        snapshot: SessionPlanSnapshot,
        output_path: Path,
        meeting_start_date_str: str,
        meeting_tz_offset_min: int = 120,
        only_tracked: bool = False,
        only_drafting: bool = False,
        tracked_tokens: Optional[Set[str]] = None,
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
    ) -> Path:
        """
        Exports an RFC 5545 .ics Calendar file strictly anchored in UTC.
        Guarantees correct calendar reminder alerts across all devices.
        """
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        annotations = annotations or {}
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        base_date = datetime.strptime(meeting_start_date_str, "%Y-%m-%d").date()

        lines = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            f"PRODID:-//3GPP Meeting Tools//SA2 #{snapshot.meeting_number} Session Plan//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-CALNAME:SA2 #{snapshot.meeting_number} ({snapshot.tdoc_id})",
        ]

        streams = snapshot.streams

        for s in snapshot.slots:
            if not s.raw_text.strip() or s.raw_text.strip() == "-":
                continue

            if only_drafting and not s.is_drafting:
                continue

            # Check if tracked
            is_tracked = False
            for it in s.items:
                if it.ai_number.lower() in tracked_tokens or any(tk in it.topic_label.lower() for tk in tracked_tokens):
                    is_tracked = True
                    break

            if only_tracked and not is_tracked and not s.is_drafting:
                continue

            # Calculate physical date
            from datetime import timedelta
            slot_date = base_date + timedelta(days=s.day_index)
            slot_date_str = slot_date.strftime("%Y-%m-%d")

            utc_start = SessionPlanTimezone.venue_local_to_utc_dt(
                slot_date_str, s.start_time_local, meeting_tz_offset_min
            )
            utc_end = SessionPlanTimezone.venue_local_to_utc_dt(
                slot_date_str, s.end_time_local, meeting_tz_offset_min
            )

            if not utc_start or not utc_end:
                continue

            dt_stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            dt_start = utc_start.strftime("%Y%m%dT%H%M%SZ")
            dt_end = utc_end.strftime("%Y%m%dT%H%M%SZ")
            uid = f"sa2-{snapshot.meeting_number}-{s.slot_key}-{dt_start}@3gpp.tools"

            s_info = streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"
            convener = s_info.convener if s_info else ""
            tohru = s_info.tohru_room if s_info else ""

            summary_prefix = "🛠️ [Drafting]" if s.is_drafting else f"[Str {s.stream_number}]"
            summary = f"{summary_prefix} {s.raw_text[:50]}"
            location = f"{r_name}"

            desc_lines = [
                f"Session: SA2 #{snapshot.meeting_number} ({snapshot.tdoc_id})",
                f"Room: {r_name}",
                f"Convener: {convener}",
                f"TOHRU: https://tohru.3gpp.org/ (Room: {tohru})",
                f"Scheduled: {s.raw_text}",
            ]

            note_obj = annotations.get(s.slot_key)
            if note_obj and note_obj.comment:
                desc_lines.append(f"My Notes: {note_obj.comment}")

            description = "\\n".join(desc_lines)

            lines.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{dt_stamp}",
                f"DTSTART:{dt_start}",
                f"DTEND:{dt_end}",
                f"SUMMARY:{summary}",
                f"LOCATION:{location}",
                f"DESCRIPTION:{description}",
                "STATUS:CONFIRMED",
                "END:VEVENT",
            ])

        lines.append("END:VCALENDAR")
        output_path.write_text("\r\n".join(lines), encoding="utf-8")
        logger.info(f"[SessionPlanExporter] Generated .ics calendar: {output_path}")
        return output_path

    @classmethod
    def generate_email_summary(
        cls,
        snapshot: SessionPlanSnapshot,
        meeting_tz_offset_min: int,
        tz_name: str,
        diff_result: Optional[SessionPlanDiffResult] = None,
        tracked_tokens: Optional[Set[str]] = None,
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
    ) -> Tuple[str, str]:
        """
        Generates pre-formatted rich HTML and plain text email bodies ready for Outlook.
        Returns: (subject, html_body)
        """
        annotations = annotations or {}
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        subject = f"[SA2#{snapshot.meeting_number}] Session Plan Update: {snapshot.tdoc_id}"

        # 1. Changes
        changes_html = ""
        if diff_result and diff_result.has_changes:
            bullets = "".join(f"<li>{html.escape(b)}</li>" for b in diff_result.changelog_bullets[:12])
            changes_html = f"""
            <div style="background:#F0FDF4; border:1px solid #BBF7D0; border-radius:6px; padding:10px 14px; margin-bottom:14px;">
              <h3 style="margin-top:0; color:#166534; font-size:14px;">🔄 Key Changes from Previous Plan ({diff_result.old_snapshot_label}):</h3>
              <ul style="margin-bottom:0; padding-left:18px; font-size:13px; color:#14532D;">
                {bullets}
              </ul>
            </div>
            """

        # 2. Drafting sessions
        drafting_slots = [s for s in snapshot.slots if s.is_drafting]
        drafting_rows = ""
        for s in drafting_slots:
            time_str = SessionPlanTimezone.format_dual_time_range(s.start_time_local, s.end_time_local, meeting_tz_offset_min)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"
            drafting_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

        drafting_html = f"""
        <h3 style="color:#991B1B; font-size:14px; margin-bottom:6px;">🛠️ Upcoming Drafting Sessions:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#FEF2F2; color:#991B1B;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue / Vienna)</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
          </tr>
          {drafting_rows}
        </table>
        """ if drafting_slots else ""

        # 3. Tracked topics
        tracked_slots = []
        for s in snapshot.slots:
            for it in s.items:
                if it.ai_number.lower() in tracked_tokens or any(tk in it.topic_label.lower() for tk in tracked_tokens):
                    tracked_slots.append(s)
                    break

        tracked_rows = ""
        for s in tracked_slots:
            time_str = SessionPlanTimezone.format_dual_time_range(s.start_time_local, s.end_time_local, meeting_tz_offset_min)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"
            tracked_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

        tracked_html = f"""
        <h3 style="color:#5B21B6; font-size:14px; margin-bottom:6px;">🎯 Our Tracked Topics Schedule:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#F5F3FF; color:#5B21B6;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue / Vienna)</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
          </tr>
          {tracked_rows}
        </table>
        """ if tracked_slots else ""

        html_body = f"""
        <div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#1E293B; line-height:1.5;">
          <p>Hi Team,</p>
          <p>Here is the latest session plan update (<b>{snapshot.tdoc_id}</b>) for <b>SA2 #{snapshot.meeting_number}</b>.</p>
          <p style="font-size:12px; color:#64748B;">📍 Venue Timezone: <b>{tz_name}</b></p>
          {changes_html}
          {drafting_html}
          {tracked_html}
          <p style="font-size:12px; color:#64748B;"><i>Generated via 3GPP Meeting Tools.</i></p>
        </div>
        """

        return subject, html_body
