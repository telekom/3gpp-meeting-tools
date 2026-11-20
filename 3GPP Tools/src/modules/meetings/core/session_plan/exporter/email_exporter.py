# --- File: src/modules/meetings/core/session_plan/exporter/email_exporter.py ---
"""
E-mail summary generators and delegation plan JSON serialization/parsing for Outlook paste.
"""
import html
import json
import re
from typing import Dict, Optional, Set, Tuple

from modules.meetings.core.session_plan.exporter.helpers import (
    get_slot_effective_delegates,
    is_slot_tracked,
)
from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffResult
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    UserSessionAnnotation,
    format_venue_time_range,
)


def generate_email_summary(
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
        time_str = format_venue_time_range(s.start_time_local, s.end_time_local)
        s_info = snapshot.streams.get(s.stream_number)
        r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
        drafting_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

    drafting_html = f"""
        <h3 style="color:#991B1B; font-size:14px; margin-bottom:6px;">🛠️ Upcoming Drafting Sessions:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#FEF2F2; color:#991B1B;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue)</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
          </tr>
          {drafting_rows}
        </table>
        """ if drafting_slots else ""

    # 3. Tracked topics
    tracked_slots = [s for s in snapshot.slots if is_slot_tracked(s, tracked_tokens)]

    tracked_rows = ""
    for s in tracked_slots:
        time_str = format_venue_time_range(s.start_time_local, s.end_time_local)
        s_info = snapshot.streams.get(s.stream_number)
        r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
        tracked_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

    tracked_html = f"""
        <h3 style="color:#5B21B6; font-size:14px; margin-bottom:6px;">🎯 Our Tracked Topics Schedule:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#F5F3FF; color:#5B21B6;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue)</th>
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


# --- Delegation Plan Serialization & Email Generation ---

def serialize_delegation_plan_json(plan: DelegationPlan) -> str:
    """Serializes a DelegationPlan to formatted JSON string."""
    data = {
        "3gpp_delegation_plan_version": 1,
        "meeting_number": str(plan.meeting_number),
        "members": [
            {
                "name": m.name,
                "email": m.email,
                "color": m.color,
                "is_self": m.is_self,
                "notes": m.notes,
            }
            for m in plan.members.values()
        ],
        "topic_owners": {
            ai: {"owner": to.owner_name, "notes": to.notes}
            for ai, to in plan.topic_owners.items()
        },
        "appointments": {
            sk: {
                "appointed_delegate": ap.appointed_delegate,
                "original_owner": ap.original_owner,
                "reason": ap.reason,
            }
            for sk, ap in plan.appointments.items()
        },
    }
    return json.dumps(data, indent=2, ensure_ascii=False)


def parse_delegation_plan_json(raw_text: str) -> Optional[dict]:
    """
    Parses delegation JSON from raw string, file content, or email text with embedded comment.
    Looks for <!-- 3GPP-DELEGATION-PLAN: {...} --> or top-level JSON.
    """
    if not raw_text or not raw_text.strip():
        return None

    # 1. Look for embedded HTML comment payload
    match = re.search(r"<!--\s*3GPP-DELEGATION-PLAN:\s*(\{.+?\})\s*-->", raw_text, re.DOTALL)
    if match:
        try:
            data = json.loads(match.group(1))
            if isinstance(data, dict) and "members" in data:
                return data
        except Exception:
            pass

    # 2. Look for JSON markdown block or raw JSON object
    json_match = re.search(r"```(?:json)?\s*(\{.+?\})\s*```", raw_text, re.DOTALL)
    candidate = json_match.group(1) if json_match else raw_text.strip()

    try:
        data = json.loads(candidate)
        if isinstance(data, dict) and ("members" in data or "topic_owners" in data):
            return data
    except Exception:
        pass

    return None


def generate_delegation_email_summary(
    snapshot: SessionPlanSnapshot,
    plan: DelegationPlan,
    meeting_tz_offset_min: int,
    tz_name: str,
    agenda_map: Optional[dict] = None,
) -> Tuple[str, str, str]:
    """
    Generates executive delegation email summary with:
    1. Roster & Workload Breakdown (Color badges)
    2. Master Topic Ownership Table
    3. Ad-Hoc Session Coverages (Collision Appointments)
    4. Individual Per-Delegate Daily Agendas
    5. Embedded machine-readable JSON payload for 1-click import.
    Returns: (subject, html_body, plain_text)
    """
    agenda_map = agenda_map or {}
    members = plan.members
    topic_owners = plan.topic_owners
    appointments = plan.appointments

    subject = f"[SA2#{snapshot.meeting_number}] Delegation Plan & AI Distribution ({snapshot.tdoc_id})"

    # Compute effective assignments for each slot in snapshot: slot_key -> List of delegate names
    slot_delegate_map: Dict[str, list] = {}
    for s in snapshot.slots:
        dels, _, is_all = get_slot_effective_delegates(s, plan, agenda_map=agenda_map)
        slot_delegate_map[s.slot_key] = dels

    # Calculate quarters workload per member
    member_workload: Dict[str, int] = {name: 0 for name in members}
    for s in snapshot.slots:
        for del_name in slot_delegate_map.get(s.slot_key, []):
            if del_name in member_workload:
                member_workload[del_name] += 1

    # 1. Executive Roster Table
    roster_rows = ""
    plain_roster = []
    for name, m in members.items():
        wl = member_workload.get(name, 0)
        owned_count = sum(1 for to in topic_owners.values() if name in to.get_owners(list(members.keys())))
        cov_count = sum(1 for ap in appointments.values() if name in ap.get_appointed_delegates(list(members.keys())))
        role_str = m.notes or "Delegate"
        roster_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><span style='display:inline-block; width:10px; height:10px; background-color:{m.color}; border-radius:50%; margin-right:6px;'></span><b>{html.escape(name)}</b></td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{html.escape(role_str)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'>{owned_count} AIs</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'>{cov_count} slots</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'><b>{wl}</b> slots</td>
            </tr>"""
        plain_roster.append(f"- {name} ({role_str}): {owned_count} AIs owned, {wl} slots scheduled")

    # 2. Topic Ownership Table
    topic_rows = ""
    plain_topics = []
    for ai_num in sorted(topic_owners.keys()):
        to = topic_owners[ai_num]
        ai_info = agenda_map.get(ai_num)
        desc = ai_info.description if ai_info else f"Agenda Item {ai_num}"
        if to.is_not_followed():
            badge_html = "<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:#F1F5F9; border:1px solid #CBD5E1; color:#64748B; font-weight:600;'>⛔ Not Followed</span>"
            plain_owner = "Not Followed"
        elif to.is_all():
            badge_html = "<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:#EEF2FF; border:1px solid #C7D2FE; color:#4338CA; font-weight:600;'>👥 All Team</span>"
            plain_owner = "All Team"
        else:
            assigned = to.get_owners(list(members.keys()))
            badges = []
            for oname in assigned:
                m = members.get(oname)
                col = m.color if m else "#2563EB"
                bg_col = m.bg_tint if m else "#E4ECFC"
                bdr_col = m.border_tint if m else "#B2C8F8"
                badges.append(f"<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:{bg_col}; border:1px solid {bdr_col}; color:{col}; font-weight:600; margin-right:4px;'>● {html.escape(oname)}</span>")
            badge_html = "".join(badges) or "—"
            plain_owner = ", ".join(assigned) or "Unassigned"

        topic_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><b>{html.escape(ai_num)}</b></td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{html.escape(desc)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{badge_html}</td>
            </tr>"""
        plain_topics.append(f"- AI {ai_num}: {desc} -> {plain_owner}")

    # 3. Ad-Hoc Session Coverages (Appointments)
    appt_rows = ""
    plain_appts = []
    for sk in sorted(appointments.keys()):
        ap = appointments[sk]
        parts = sk.split("|")
        day_name = parts[0] if len(parts) > 0 else ""
        q_name = parts[1] if len(parts) > 1 else ""
        str_num = parts[3] if len(parts) > 3 else "1"
        orig_txt = f"for {ap.original_owner}" if ap.original_owner else ""
        reason_txt = f" ({ap.reason})" if ap.reason else ""

        if ap.is_all():
            badge_app = "<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:#EEF2FF; border:1px solid #C7D2FE; color:#4338CA; font-weight:600;'>👥 All Team</span>"
            plain_app = "All Team"
        else:
            assigned = ap.get_appointed_delegates(list(members.keys()))
            badges = []
            for oname in assigned:
                m_app = members.get(oname)
                col_app = m_app.color if m_app else "#047857"
                bg_app = m_app.bg_tint if m_app else "#E0EEEA"
                bdr_app = m_app.border_tint if m_app else "#A7CFC4"
                badges.append(f"<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:{bg_app}; border:1px solid {bdr_app}; color:{col_app}; font-weight:600; margin-right:4px;'>● {html.escape(oname)}</span>")
            badge_app = "".join(badges) or "—"
            plain_app = ", ".join(assigned) or "Unassigned"

        appt_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><b>{day_name}</b> {q_name} (Str {str_num})</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{badge_app} {html.escape(orig_txt)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; color:#64748B;'>{html.escape(ap.reason or 'Ad-hoc cover')}</td>
            </tr>"""
        plain_appts.append(f"- {day_name} {q_name} (Stream {str_num}): {plain_app} covering {orig_txt}{reason_txt}")

    appt_html = f"""
        <div style="margin-bottom:20px;">
          <h3 style="color:#B45309; font-size:14px; margin-bottom:8px;">⚡ Ad-Hoc Session Coverages & Clash Resolutions:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:12px;">
            <tr style="background:#FFFBEB; color:#92400E;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Covering Delegate</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Reason</th>
            </tr>
            {appt_rows}
          </table>
        </div>
        """ if appointments else ""

    # 4. Individual Timetables (per delegate)
    individual_sections_html = ""
    for name, m in members.items():
        del_slots = [s for s in snapshot.slots if name in slot_delegate_map.get(s.slot_key, [])]
        if not del_slots:
            continue

        del_rows = ""
        for s in del_slots:
            time_str = format_venue_time_range(s.start_time_local, s.end_time_local)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
            is_cov = (
                s.slot_key in appointments
                and name in appointments[s.slot_key].get_appointed_delegates(list(members.keys()))
            )
            role_badge = "<span style='background:#FEF3C7; color:#92400E; padding:2px 6px; border-radius:4px; font-size:10px;'>Coverage</span>" if is_cov else "<span style='background:#EFF6FF; color:#1E40AF; padding:2px 6px; border-radius:4px; font-size:10px;'>Owner</span>"

            del_rows += f"""<tr>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>{time_str}</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>Str {s.stream_number} ({html.escape(r_name)})</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>{role_badge} {html.escape(s.raw_text)}</td>
                </tr>"""

        individual_sections_html += f"""
            <div style="margin-bottom:16px;">
              <h4 style="color:{m.color}; margin-bottom:6px; font-size:13px;">📅 {html.escape(name)}'s Schedule ({len(del_slots)} sessions):</h4>
              <table style="border-collapse:collapse; width:100%; font-size:11px; margin-bottom:8px;">
                <tr style="background:#F8FAFC;">
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue)</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
                </tr>
                {del_rows}
              </table>
            </div>
            """

    # 5. Embedded Machine-Readable JSON Payload
    json_payload = serialize_delegation_plan_json(plan)
    embedded_comment = f"\n<!-- 3GPP-DELEGATION-PLAN: {json_payload} -->\n"

    html_body = f"""
        <div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#1E293B; line-height:1.5;">
          <p>Hi Team,</p>
          <p>Here is the delegation plan and topic allocation for <b>SA2 #{snapshot.meeting_number}</b> (Session Plan: <b>{snapshot.tdoc_id}</b>).</p>
          <p style="font-size:12px; color:#64748B;">📍 Venue Timezone: <b>{tz_name}</b></p>

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">👥 Delegation Roster & Workload:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
            <tr style="background:#F1F5F9; color:#334155;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Delegate</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Role</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Owned Topics</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Coverages</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Total Quarters</th>
            </tr>
            {roster_rows}
          </table>

          {appt_html}

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">📋 Master Topic Ownership:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
            <tr style="background:#F1F5F9; color:#334155;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">AI</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Lead Owner</th>
            </tr>
            {topic_rows}
          </table>

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">📅 Individual Delegate Schedules:</h3>
          {individual_sections_html}

          <p style="font-size:11px; color:#94A3B8; margin-top:20px;">
            <i>Tip: To import this plan into your 3GPP Tools, copy this email text and click <b>Import Team Plan</b> in the Session Plan window.</i>
          </p>
        </div>
        {embedded_comment}
        """

    plain_text = f"SA2 #{snapshot.meeting_number} Delegation Plan ({snapshot.tdoc_id})\n"
    plain_text += f"Venue Timezone: {tz_name}\n\n"
    plain_text += "--- ROSTER ---\n" + "\n".join(plain_roster) + "\n\n"
    if plain_appts:
        plain_text += "--- AD-HOC COVERAGES ---\n" + "\n".join(plain_appts) + "\n\n"
    plain_text += "--- TOPIC OWNERSHIP ---\n" + "\n".join(plain_topics) + "\n"

    return subject, html_body, plain_text
