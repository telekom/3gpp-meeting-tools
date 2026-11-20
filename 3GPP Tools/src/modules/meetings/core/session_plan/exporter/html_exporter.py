# --- File: src/modules/meetings/core/session_plan/exporter/html_exporter.py ---
"""
Document synthesis and entry point for standalone offline mobile HTML Session Plan export.
Combines prepared context, Overview Grid, Mobile Day Cards, CSS, and interactive JS.
"""
import html
import json
import re
from pathlib import Path
from typing import Dict, Optional, Set

from modules.meetings.core.session_plan.exporter.css_grid import CSS_GRID
from modules.meetings.core.session_plan.exporter.css_layout import CSS_LAYOUT
from modules.meetings.core.session_plan.exporter.css_views import CSS_VIEWS
from modules.meetings.core.session_plan.exporter.helpers import logger
from modules.meetings.core.session_plan.exporter.html_cards import build_mobile_cards_html
from modules.meetings.core.session_plan.exporter.html_data import build_mobile_html_context
from modules.meetings.core.session_plan.exporter.html_overview import (
    build_delegate_filter_buttons_html,
    build_overview_grid_html,
)
from modules.meetings.core.session_plan.exporter.js_conflicts_unassigned import JS_CONFLICTS_UNASSIGNED
from modules.meetings.core.session_plan.exporter.js_core import JS_CORE
from modules.meetings.core.session_plan.exporter.js_now_live import JS_NOW_LIVE
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    UserSessionAnnotation,
)


def export_mobile_html(
    snapshot: SessionPlanSnapshot,
    output_path: Path,
    meeting_tz_offset_min: int = 120,
    tz_name: str = "Venue Time",
    meeting_dates_str: str = "",
    annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
    tracked_tokens: Optional[Set[str]] = None,
    delegation_plan: Optional[DelegationPlan] = None,
    agenda_map: Optional[dict] = None,
    room_locations: Optional[Dict[str, str]] = None,
    slot_locations: Optional[Dict[str, str]] = None,
) -> Path:
    """
    Exports a self-contained, responsive, 100% offline HTML view featuring:
    1. 🗓️ 5-Day Spatial Overview Grid (desktop-aligned rounded cards)
    2. 🔴 Now View (live/upcoming parallel streams, rooms & hand-made locations)
    3. 📱 Mobile Day Cards (day-by-day touch-friendly quarter view)
    4. ⚠️ Conflicts & ❓ Unassigned AIs views
    No external CDN dependencies (fonts and CSS are purely embedded).
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    ctx = build_mobile_html_context(
        snapshot=snapshot,
        output_path=output_path,
        meeting_tz_offset_min=meeting_tz_offset_min,
        tz_name=tz_name,
        meeting_dates_str=meeting_dates_str,
        annotations=annotations,
        tracked_tokens=tracked_tokens,
        delegation_plan=delegation_plan,
        agenda_map=agenda_map,
        room_locations=room_locations,
        slot_locations=slot_locations,
    )

    overview_html_str = build_overview_grid_html(ctx)
    cards_html_str = build_mobile_cards_html(ctx)
    del_filter_html = build_delegate_filter_buttons_html(ctx)

    css_style = CSS_LAYOUT + CSS_GRID + CSS_VIEWS
    js_script = JS_CORE + JS_CONFLICTS_UNASSIGNED + JS_NOW_LIVE

    raw_snap_lbl = (snapshot.snapshot_label or "").strip()
    short_snap_lbl = raw_snap_lbl
    if raw_snap_lbl:
        m_snap = re.match(r"^([A-Za-z0-9-]+)\s*\((.+)\)$", raw_snap_lbl)
        if m_snap:
            inner_s = re.sub(r"\.(docx?|docm|zip|htm|html)$", "", m_snap.group(2).strip(), flags=re.I)
            inner_s = re.sub(rf"^{re.escape(m_snap.group(1).strip())}[_\s-]*", "", inner_s, flags=re.I)
            inner_s = re.sub(r"^(?:SA\d+-?\d+[-\s]*)?Session\s*Plan\s*[-\s]*", "", inner_s, flags=re.I).strip(" -_")
            short_snap_lbl = inner_s or m_snap.group(1).strip()
    snap_label_html = (
        f'<span class="badge-snap" title="{html.escape(raw_snap_lbl)}">{html.escape(short_snap_lbl)}</span>'
        if raw_snap_lbl
        else ""
    )
    tz_diff_note = f" | Laptop Diff: {ctx.tz_diff_str}" if ctx.is_tz_different else ""

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>SA2 #{snapshot.meeting_number} Session Plan ({snapshot.tdoc_id})</title>
<style>
{css_style}
</style>
</head>
<body>

<header>
  <div class="header-top">
    <div class="title-block">
      <h1 class="title">SA2 #{snapshot.meeting_number} Session Plan</h1>
      <span class="badge-tdoc">{snapshot.tdoc_id}</span>
      {snap_label_html}
      <span id="liveClockPill" class="badge-live-clock" title="📍 {html.escape(tz_name)}{html.escape(tz_diff_note)}">🏛️ Venue: <strong>--:--</strong> <span style="color:#93C5FD;">|</span> 🏠 Local: <strong>--:--</strong></span>
      <span class="badge-utc" title="📍 {html.escape(tz_name)}{html.escape(tz_diff_note)}">🕒 {ctx.venue_utc_str}</span>
    </div>
    <div class="header-actions">
      <div class="view-switcher">
        <button id="btnViewGrid" class="view-btn active" onclick="switchView('grid')">🗓️ Overview Grid</button>
        <button id="btnViewNow" class="view-btn" onclick="switchView('now')">🔴 Now</button>
        <button id="btnViewCards" class="view-btn" onclick="switchView('cards')">📱 Day Cards</button>
        <button id="btnViewConflicts" class="view-btn" onclick="switchView('conflicts')">⚠️ Conflicts (0)</button>
        <button id="btnViewUnassigned" class="view-btn" onclick="switchView('unassigned')">❓ Unassigned ({ctx.unassigned_ai_count})</button>
      </div>
      <span id="clashWarningBadge" class="badge-clash" style="display:none;" onclick="switchView('conflicts')" title="Click to open the Parallel Stream Conflicts & Coverage Board"></span>
      <button id="btnDownloadIcs" class="btn-cal-download" onclick="downloadSelectedCalendar()" title="Download RFC 5545 calendar for selected delegate or whole team">📅 Team .ics</button>
    </div>
  </div>

  <div class="controls-bar">
    <div class="del-strip">
      <span class="strip-label">Focus:</span>
      <button class="del-filter-btn active" data-del="ALL" onclick="selectDelegate('ALL')">👥 All Team</button>
      {del_filter_html}
      <span class="strip-sep"></span>
      <button id="btnUnassigned" class="del-filter-btn btn-toggle" onclick="toggleUnassigned()" title="Highlight scheduled sessions that have no assigned delegate or contain unassigned Agenda Items">❓ Unassigned ({ctx.unassigned_slot_count})</button>
      <button id="btnDrafting" class="del-filter-btn btn-toggle" onclick="toggleDrafting()">🛠️ Drafting Only</button>
    </div>

    <div class="filter-controls-row">
      <div id="dayFilterStrip" class="day-filter-strip" title="Focus on specific meeting days or hide past days">
        <span class="strip-label">Days:</span>
        <button class="day-filter-btn active" data-day-filter="ALL" onclick="selectDayFilter('ALL')">All</button>
        <button class="day-filter-btn" data-day-filter="TODAY_PLUS" onclick="selectDayFilter('TODAY_PLUS')">Today+</button>
        <button class="day-filter-btn" data-day-filter="0" onclick="selectDayFilter(0)">Mon</button>
        <button class="day-filter-btn" data-day-filter="1" onclick="selectDayFilter(1)">Tue</button>
        <button class="day-filter-btn" data-day-filter="2" onclick="selectDayFilter(2)">Wed</button>
        <button class="day-filter-btn" data-day-filter="3" onclick="selectDayFilter(3)">Thu</button>
        <button class="day-filter-btn" data-day-filter="4" onclick="selectDayFilter(4)">Fri</button>
      </div>
      <div id="gridControls" class="grid-row-controls">
        <label class="toggle-label"><input type="checkbox" id="chkStr1" checked onchange="toggleStream(1)"><span class="str-dot s1"></span> Str 1</label>
        <label class="toggle-label"><input type="checkbox" id="chkStr2" checked onchange="toggleStream(2)"><span class="str-dot s2"></span> Str 2</label>
        <label class="toggle-label"><input type="checkbox" id="chkStr3" checked onchange="toggleStream(3)"><span class="str-dot s3"></span> Str 3</label>
        <label class="toggle-label"><input type="checkbox" id="chkBreaks" checked onchange="toggleBreaks()"> Breaks</label>
      </div>
      <input type="text" id="searchInput" class="search-box" placeholder="🔍 Filter AI, topic, room..." oninput="applyFilters()">
    </div>
  </div>
</header>

<!-- Main View 1: 5-Day Spatial Overview Grid -->
<div id="overviewView">
  <div id="liveVenueBanner" class="live-banner" style="display:none;margin-bottom:5px;" onclick="switchView('now')"></div>
  {overview_html_str}
</div>

<!-- Main View 2: Live "Now" & Upcoming Rooms Dashboard -->
<div id="nowView" style="display:none;">
  <div class="now-view-shell">
    <div id="nowViewContent"></div>
    <div id="nowDispatchPanel" class="now-dispatch-panel" style="display:none;margin-top:12px;"></div>
  </div>
</div>

<!-- Main View 3: Responsive Mobile Day Cards -->
<div id="cardsView" style="display:none;">
  {cards_html_str}
</div>

<!-- Main View 3: Parallel Stream Conflicts & Coverage Board -->
<div id="conflictsView" style="display:none;">
  <div class="aux-view-container">
    <div class="aux-header-card">
      <div>
        <h2 class="aux-title">⚠️ Parallel Stream Conflicts & Coverage Board</h2>
        <div class="aux-subtitle">Detects double-booked delegates across parallel rooms (with half-quarter TU 1 / TU 2 precision) and identifies free colleagues who can cover.</div>
      </div>
    </div>
    <div id="conflictsViewContent"></div>
  </div>
</div>

<!-- Main View 4: Unassigned Agenda Items & Uncovered Sessions -->
<div id="unassignedView" style="display:none;">
  <div class="aux-view-container">
    <div class="aux-header-card">
      <div>
        <h2 class="aux-title">❓ Unassigned Agenda Items & Uncovered Sessions</h2>
        <div class="aux-subtitle">Sessions and Agenda Items that do not yet have a default AI owner or session-level delegate coverage.</div>
      </div>
    </div>
    <div id="unassignedViewContent"></div>
  </div>
</div>

<!-- Click-Activated Slide-In Slot Inspector Drawer -->
<aside id="slotInspectorDrawer" class="slot-inspector-drawer" aria-hidden="true" role="dialog" aria-label="Session Details">
  <div class="drawer-header">
    <div class="drawer-heading-group">
      <span class="drawer-kicker">Session Details</span>
      <span class="drawer-hint">Click another slot to switch · Esc to close</span>
    </div>
    <button type="button" class="drawer-close-btn" onclick="closeSlotInspector()" aria-label="Close session details" title="Close (Esc)">✕</button>
  </div>
  <div id="slotInspectorBody" class="drawer-body"></div>
</aside>

<script>
  const CAL_EVENTS = {ctx.cal_events_json};
  const ROSTER_MEMBERS = {ctx.roster_members_json};
  const ROOM_LOCATIONS = {ctx.room_locations_json};
  const STREAMS_META = {ctx.streams_meta_json};
  const UNASSIGNED_AIS = {ctx.unassigned_ais_json};
  const UNASSIGNED_SLOTS = {ctx.unassigned_slots_json};
  const MEETING_BASE_DATE = {json.dumps(ctx.base_date_js)};
  const MEETING_TZ_OFFSET_MIN = {meeting_tz_offset_min};
  const MEETING_NUMBER = {json.dumps(str(snapshot.meeting_number or ""))};
  const TDOC_ID = {json.dumps(str(snapshot.tdoc_id or ""))};
</script>
<script>
{js_script}
</script>
</body>
</html>
"""

    output_path.write_text(full_html, encoding="utf-8")
    logger.info(f"[SessionPlanExporter] Generated interactive Session Plan HTML: {output_path} ({output_path.stat().st_size} bytes)")
    return output_path
