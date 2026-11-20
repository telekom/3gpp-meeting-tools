# --- File: src/modules/meetings/core/session_plan/exporter/js_conflicts_unassigned.py ---
"""
Embedded JS (part 2/3): Parallel-stream clash detection, Conflicts board
and Unassigned AIs view rendering.
"""

JS_CONFLICTS_UNASSIGNED = """\
  // ------------------------------------------------------------------
  // Conflict & Parallel Stream Clash Detection + Conflicts View Tab
  // ------------------------------------------------------------------
  function getAllRosterNames() {
    if (typeof ROSTER_MEMBERS !== 'undefined' && Array.isArray(ROSTER_MEMBERS) && ROSTER_MEMBERS.length > 0) {
      return ROSTER_MEMBERS.map(m => m.name);
    }
    const allMembers = [];
    document.querySelectorAll('.del-filter-btn[data-del]').forEach(b => {
      const d = b.getAttribute('data-del');
      if (d && d !== 'ALL') allMembers.push(d);
    });
    return allMembers;
  }

  function getRosterMemberMeta(name) {
    if (typeof ROSTER_MEMBERS !== 'undefined' && Array.isArray(ROSTER_MEMBERS)) {
      const found = ROSTER_MEMBERS.find(m => m.name.toLowerCase() === String(name || '').toLowerCase());
      if (found) return found;
    }
    return { name: name, initials: String(name || '').substring(0, 2).toUpperCase(), color: '#2563EB', bg_tint: '#EFF6FF', border_tint: '#BFDBFE' };
  }

  function getFreeMembersForSlot(dayIdx, quarter, subSlot) {
    const tus = (subSlot === 1) ? [1] : (subSlot === 2) ? [2] : [1, 2];
    const busyLower = new Set();
    for (const ev of CAL_EVENTS) {
      if (ev.day_idx !== dayIdx || ev.quarter !== quarter) continue;
      const evTus = (ev.sub_slot === 1) ? [1] : (ev.sub_slot === 2) ? [2] : [1, 2];
      const overlaps = tus.some(t => evTus.includes(t));
      if (!overlaps) continue;
      if (ev.is_all) {
        return []; // Whole team is busy
      }
      (ev.delegates || []).forEach(d => busyLower.add(d.toLowerCase()));
    }
    return getAllRosterNames().filter(n => !busyLower.has(n.toLowerCase()));
  }

  function buildConflictGroups(filterMember) {
    const membersToCheck = (filterMember && filterMember !== 'ALL') ? [filterMember] : getAllRosterNames();
    const groups = [];
    for (const member of membersToCheck) {
      const mLower = member.toLowerCase();
      const slotGroups = {};
      for (const ev of CAL_EVENTS) {
        const effDels = (ev.delegates || []).map(d => d.toLowerCase());
        if (!effDels.includes(mLower) && !ev.is_all) continue;
        const k = ev.day_idx + '|' + ev.quarter;
        if (!slotGroups[k]) slotGroups[k] = [];
        slotGroups[k].push(ev);
      }
      for (const k in slotGroups) {
        const evList = slotGroups[k];
        if (evList.length < 2) continue;
        // Check TU overlap across distinct streams
        for (const tu of [1, 2]) {
          const activeInTu = evList.filter(e => e.sub_slot === 0 || e.sub_slot === tu);
          const distinctStreams = new Set(activeInTu.map(e => e.stream_num));
          if (distinctStreams.size > 1) {
            const sortedEvs = activeInTu.slice().sort((a, b) => a.stream_num - b.stream_num);
            const sig = member + '|' + k + '|' + sortedEvs.map(e => e.stream_num + ':' + e.slot_key).join(',');
            const existing = groups.find(g => g.sig === sig);
            if (existing) {
              if (!existing.tus.includes(tu)) existing.tus.push(tu);
            } else {
              groups.push({
                sig: sig,
                member: member,
                day_idx: sortedEvs[0].day_idx,
                day_name: sortedEvs[0].day_name || DAY_NAMES_LIST[sortedEvs[0].day_idx] || ('Day ' + (sortedEvs[0].day_idx + 1)),
                quarter: sortedEvs[0].quarter,
                tus: [tu],
                events: sortedEvs
              });
            }
          }
        }
      }
    }
    groups.sort((a, b) => (a.day_idx - b.day_idx) || (PERIOD_LIST.indexOf(a.quarter) - PERIOD_LIST.indexOf(b.quarter)) || a.member.localeCompare(b.member));
    return groups;
  }

  function detectDelegateClashes(delName) {
    const clashBadge = document.getElementById('clashWarningBadge');
    const clashTabBtn = document.getElementById('btnViewConflicts');

    document.querySelectorAll('.has-clash').forEach(el => el.classList.remove('has-clash'));
    document.querySelectorAll('.clash-pill-ring').forEach(el => el.classList.remove('clash-pill-ring'));

    const allGroups = buildConflictGroups('ALL');
    if (clashTabBtn) {
      clashTabBtn.textContent = '⚠️ Conflicts (' + allGroups.length + ')';
      clashTabBtn.classList.toggle('has-alert', allGroups.length > 0);
    }

    if (!clashBadge) return;

    if (delName === 'ALL') {
      const memberClashes = {};
      let totalClashes = 0;
      const allMembers = getAllRosterNames();

      for (const member of allMembers) {
        const c = getClashesForMember(member);
        if (c.length > 0) {
          memberClashes[member] = c.length;
          totalClashes += c.length;
          c.forEach(ev => markClashCells(ev, member));
        }
      }

      if (totalClashes > 0) {
        const summary = Object.entries(memberClashes).map(([m, n]) => m + ' (' + n + ')').join(', ');
        clashBadge.innerHTML = '⚠️ ' + allGroups.length + ' Conflict' + (allGroups.length > 1 ? 's' : '') + ': ' + summary;
        clashBadge.style.display = 'inline-flex';
      } else {
        clashBadge.style.display = 'none';
      }
      return;
    }

    const clashes = getClashesForMember(delName);
    const delGroups = buildConflictGroups(delName);
    if (clashes.length > 0) {
      clashBadge.innerHTML = '⚠️ ' + delGroups.length + ' Parallel Clash' + (delGroups.length > 1 ? 'es' : '') + ' for ' + delName;
      clashBadge.style.display = 'inline-flex';
      clashes.forEach(ev => markClashCells(ev, delName));
    } else {
      clashBadge.style.display = 'none';
    }
  }

  function getClashesForMember(memberName) {
    const mLower = memberName.toLowerCase();
    const tuOccupancy = {};

    for (const ev of CAL_EVENTS) {
      const effDels = (ev.delegates || []).map(d => d.toLowerCase());
      if (!effDels.includes(mLower)) continue;

      const tus = (ev.sub_slot === 1) ? [1] : (ev.sub_slot === 2) ? [2] : [1, 2];
      for (const tu of tus) {
        const key = ev.day_idx + '|' + ev.quarter + '|' + tu;
        if (!tuOccupancy[key]) tuOccupancy[key] = [];
        tuOccupancy[key].push(ev);
      }
    }

    const clashingEvents = new Set();
    for (const key in tuOccupancy) {
      const list = tuOccupancy[key];
      const streams = new Set(list.map(e => e.stream_num));
      if (streams.size > 1) {
        list.forEach(e => clashingEvents.add(e));
      }
    }
    return Array.from(clashingEvents);
  }

  function markClashCells(ev, memberName) {
    const sDay = String(ev.day_idx);
    const sPeriod = String(ev.quarter);
    const sStr = String(ev.stream_num);
    const sTu = ev.sub_slot ? String(ev.sub_slot) : null;
    const gridSel = '.grid-cell[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]';
    const cardSel = '.cards-slot-row[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]';
    const mLower = memberName ? String(memberName).toLowerCase() : '';

    const applyClashUi = el => {
      el.classList.add('has-clash');
      if (mLower) {
        el.querySelectorAll('.del-pill-badge[data-del-name]').forEach(pill => {
          if ((pill.getAttribute('data-del-name') || '').toLowerCase() === mLower) {
            pill.classList.add('clash-pill-ring');
          }
        });
      }
    };

    if (sTu && sTu !== '0') {
      const tuGrid = document.querySelectorAll(gridSel + '[data-tu="' + sTu + '"]');
      if (tuGrid.length > 0) {
        tuGrid.forEach(applyClashUi);
      } else {
        document.querySelectorAll(gridSel).forEach(applyClashUi);
      }
      const tuCard = document.querySelectorAll(cardSel + '[data-tu="' + sTu + '"]');
      if (tuCard.length > 0) {
        tuCard.forEach(applyClashUi);
      } else {
        document.querySelectorAll(cardSel).forEach(applyClashUi);
      }
    } else {
      document.querySelectorAll(gridSel).forEach(applyClashUi);
      document.querySelectorAll(cardSel).forEach(applyClashUi);
    }
  }

  function renderConflictsView() {
    const container = document.getElementById('conflictsViewContent');
    if (!container) return;
    const sInput = document.getElementById('searchInput');
    const query = ((sInput && sInput.value) || '').toLowerCase().trim();

    let groups = buildConflictGroups(selectedDelegate);
    if (query) {
      groups = groups.filter(g => {
        const hay = (g.member + ' ' + g.day_name + ' ' + g.quarter + ' ' + g.events.map(e => e.title + ' ' + e.room).join(' ')).toLowerCase();
        return hay.includes(query);
      });
    }

    if (groups.length === 0) {
      const scopeLbl = (selectedDelegate === 'ALL') ? 'across the delegation' : ('for ' + escHtml(selectedDelegate));
      container.innerHTML = '<div class="empty-state-box">✅ <strong>Zero parallel stream conflicts detected ' + scopeLbl + '!</strong><br><span style="font-size:12px;color:#64748B;">All assigned sessions occur in non-overlapping time units.</span></div>';
      return;
    }

    let htmlRows = '';
    for (const g of groups) {
      const mMeta = getRosterMemberMeta(g.member);
      const tuLabel = (g.tus.length === 2) ? 'Full Quarter (TU 1 & 2)' : ('TU ' + g.tus[0] + ' Only');
      const subSlotFilter = (g.tus.length === 1) ? g.tus[0] : 0;
      const freeNames = getFreeMembersForSlot(g.day_idx, g.quarter, subSlotFilter);

      let streamsHtml = '';
      for (const ev of g.events) {
        const roleTag = ev.is_all ? '👥 All Team' : (ev.is_cov ? '⚡ Session Override' : '📌 Default AI Owner');
        const tuBadge = (ev.sub_slot === 1) ? ' [TU 1]' : (ev.sub_slot === 2) ? ' [TU 2]' : '';
        streamsHtml += '<div class="clash-stream-box">' +
          '<div style="display:flex;justify-content:space-between;gap:6px;font-weight:700;font-size:11.5px;color:#991B1B;">' +
          '<span>Stream ' + ev.stream_num + tuBadge + ' · 📍 ' + escHtml(ev.room) + '</span>' +
          '<span style="font-size:10px;color:#B45309;background:#FEF3C7;padding:1px 5px;border-radius:4px;">' + roleTag + '</span>' +
          '</div>' +
          '<div style="font-size:11.5px;color:#0F172A;margin-top:2px;font-weight:600;">' + escHtml(ev.title) + '</div>' +
          '</div>';
      }

      let freeHtml = '';
      if (freeNames.length > 0) {
        freeHtml = freeNames.map(fn => {
          const fm = getRosterMemberMeta(fn);
          return '<span class="free-colleague-pill" style="background:' + fm.bg_tint + ';color:' + fm.color + ';border:1px solid ' + fm.border_tint + ';">🟢 ' + escHtml(fn) + '</span>';
        }).join(' ');
      } else {
        freeHtml = '<span style="color:#DC2626;font-size:11px;font-weight:600;">⚠️ All roster colleagues busy</span>';
      }

      htmlRows += '<tr>' +
        '<td style="white-space:nowrap;">' +
          '<span style="background:' + mMeta.bg_tint + ';color:' + mMeta.color + ';border:1px solid ' + mMeta.border_tint + ';padding:3px 9px;border-radius:999px;font-weight:700;font-size:11.5px;">● ' + escHtml(g.member) + '</span>' +
        '</td>' +
        '<td style="white-space:nowrap;">' +
          '<strong>' + escHtml(g.day_name) + ' · ' + escHtml(g.quarter) + '</strong><br>' +
          '<span style="font-size:11px;color:#64748B;">' + tuLabel + '</span>' +
        '</td>' +
        '<td>' + streamsHtml + '</td>' +
        '<td>' + freeHtml + '</td>' +
        '<td style="text-align:right;white-space:nowrap;">' +
          '<button class="btn-jump-grid" onclick="jumpToSlotInGrid(' + g.day_idx + ', \\'' + escHtml(g.quarter) + '\\', ' + g.events[0].stream_num + ')">🔍 Show in Grid</button>' +
        '</td>' +
      '</tr>';
    }

    container.innerHTML = '<table class="aux-table">' +
      '<thead><tr>' +
        '<th>Clashing Delegate</th>' +
        '<th>Day / Period / TU</th>' +
        '<th>Conflicting Parallel Streams</th>' +
        '<th>🟢 Available / Free Colleagues (Can Cover)</th>' +
        '<th style="text-align:right;">Action</th>' +
      '</tr></thead>' +
      '<tbody>' + htmlRows + '</tbody></table>';
  }

  // ------------------------------------------------------------------
  // Unassigned AIs View Tab
  // ------------------------------------------------------------------
  function toggleUnassignedScope(scheduledOnly) {
    unassignedScheduledOnly = scheduledOnly;
    renderUnassignedView();
  }

  function renderUnassignedView() {
    const container = document.getElementById('unassignedViewContent');
    if (!container) return;
    const sInput = document.getElementById('searchInput');
    const query = ((sInput && sInput.value) || '').toLowerCase().trim();

    const aiList = (typeof UNASSIGNED_AIS !== 'undefined' && Array.isArray(UNASSIGNED_AIS)) ? UNASSIGNED_AIS : [];
    const slotList = (typeof UNASSIGNED_SLOTS !== 'undefined' && Array.isArray(UNASSIGNED_SLOTS)) ? UNASSIGNED_SLOTS : [];

    // Filter unassigned scheduled slots
    let filteredSlots = slotList.filter(s => {
      if (!query) return true;
      const hay = (s.day_name + ' ' + s.quarter + ' ' + s.title + ' ' + s.room + ' ' + (s.ai_numbers || []).join(' ')).toLowerCase();
      return hay.includes(query);
    });

    // Filter unassigned AI catalog
    let filteredAis = aiList.filter(a => {
      if (unassignedScheduledOnly && a.scheduled_count === 0) return false;
      if (!query) return true;
      const hay = ('ai ' + a.ai_number + ' ' + a.title).toLowerCase();
      return hay.includes(query);
    });

    let html = '';

    // Section 1: Unassigned Scheduled Session Slots
    html += '<div style="margin-bottom:18px;">' +
      '<h3 style="font-size:13.5px;font-weight:700;color:#1E293B;margin:0 0 8px 0;">📅 Unassigned Scheduled Session Slots (' + filteredSlots.length + ')</h3>';

    if (filteredSlots.length === 0) {
      html += '<div class="empty-state-box" style="padding:18px;">✅ All scheduled session slots have at least one assigned delegate!</div>';
    } else {
      let sRows = '';
      for (const s of filteredSlots) {
        const tuLbl = (s.sub_slot === 1) ? 'TU 1' : (s.sub_slot === 2) ? 'TU 2' : 'Full';
        const unAis = (s.unassigned_ais && s.unassigned_ais.length > 0)
          ? s.unassigned_ais.map(a => '<span class="ai-chip unassigned-chip">❓ AI ' + escHtml(a) + '</span>').join(' ')
          : '<span class="del-pill-badge is-unassigned-pill">❓ Entire Slot Unassigned</span>';
        const freeNames = getFreeMembersForSlot(s.day_idx, s.quarter, s.sub_slot || 0);
        const freeStr = freeNames.length > 0
          ? freeNames.map(fn => {
              const fm = getRosterMemberMeta(fn);
              return '<span class="free-colleague-pill" style="background:' + fm.bg_tint + ';color:' + fm.color + ';border:1px solid ' + fm.border_tint + ';">🟢 ' + escHtml(fn) + '</span>';
            }).join(' ')
          : '<span style="color:#94A3B8;font-size:11px;">—</span>';

        sRows += '<tr>' +
          '<td style="white-space:nowrap;"><strong>' + escHtml(s.day_name) + ' · ' + escHtml(s.quarter) + '</strong> (' + tuLbl + ')<br><span style="font-size:11px;color:#64748B;">' + escHtml(s.start_local) + '–' + escHtml(s.end_local) + '</span></td>' +
          '<td style="white-space:nowrap;"><span class="stream-badge str-' + s.stream_num + '">Stream ' + s.stream_num + '</span><br><span style="font-size:11px;color:#64748B;">📍 ' + escHtml(s.room) + '</span></td>' +
          '<td><div style="font-weight:600;color:#0F172A;">' + escHtml(s.title) + '</div><div style="margin-top:4px;">' + unAis + '</div></td>' +
          '<td>' + freeStr + '</td>' +
          '<td style="text-align:right;white-space:nowrap;"><button class="btn-jump-grid" onclick="jumpToSlotInGrid(' + s.day_idx + ', \\'' + escHtml(s.quarter) + '\\', ' + s.stream_num + ')">🔍 Show in Grid</button></td>' +
        '</tr>';
      }
      html += '<table class="aux-table"><thead><tr><th>Day / Period</th><th>Stream & Room</th><th>Session Topic & Unassigned AIs</th><th>🟢 Available Colleagues</th><th style="text-align:right;">Action</th></tr></thead><tbody>' + sRows + '</tbody></table>';
    }
    html += '</div>';

    // Section 2: Unassigned Agenda Items (AI Catalog)
    html += '<div>' +
      '<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:8px;">' +
        '<h3 style="font-size:13.5px;font-weight:700;color:#1E293B;margin:0;">📑 Unassigned Agenda Items (' + filteredAis.length + ')</h3>' +
        '<label class="toggle-label" style="font-size:12px;"><input type="checkbox" ' + (unassignedScheduledOnly ? 'checked' : '') + ' onchange="toggleUnassignedScope(this.checked)"> Show only AIs scheduled in this session plan</label>' +
      '</div>';

    if (filteredAis.length === 0) {
      html += '<div class="empty-state-box" style="padding:18px;">✅ Zero unassigned Agenda Items matching current filter!</div>';
    } else {
      let aRows = '';
      for (const a of filteredAis) {
        let schedBadges = '<span style="color:#94A3B8;font-size:11px;">Not scheduled in grid</span>';
        let jumpBtn = '';
        if (a.slots && a.slots.length > 0) {
          schedBadges = a.slots.map(sl =>
            '<span class="ai-chip" style="cursor:pointer;margin-right:4px;" onclick="jumpToSlotInGrid(' + sl.day_idx + ', \\'' + escHtml(sl.quarter) + '\\', ' + sl.stream_num + ')">📅 ' + escHtml(sl.day_name.substring(0, 3)) + ' ' + escHtml(sl.quarter) + ' (Str ' + sl.stream_num + ')</span>'
          ).join(' ');
          const first = a.slots[0];
          jumpBtn = '<button class="btn-jump-grid" onclick="jumpToSlotInGrid(' + first.day_idx + ', \\'' + escHtml(first.quarter) + '\\', ' + first.stream_num + ')">🔍 Jump to Slot</button>';
        }
        const tdocBadge = a.tdoc_count ? '<span class="ai-chip">[' + escHtml(a.tdoc_count) + ' TDocs]</span>' : '—';
        aRows += '<tr>' +
          '<td style="white-space:nowrap;font-weight:700;color:#B45309;">❓ AI ' + escHtml(a.ai_number) + '</td>' +
          '<td style="font-weight:600;color:#0F172A;">' + escHtml(a.title) + '</td>' +
          '<td>' + tdocBadge + '</td>' +
          '<td>' + schedBadges + '</td>' +
          '<td style="text-align:right;white-space:nowrap;">' + jumpBtn + '</td>' +
        '</tr>';
      }
      html += '<table class="aux-table"><thead><tr><th>AI Number</th><th>Topic / Work Item Description</th><th>TDocs</th><th>Scheduled Sessions</th><th style="text-align:right;">Action</th></tr></thead><tbody>' + aRows + '</tbody></table>';
    }
    html += '</div>';

    container.innerHTML = html;
  }

"""
