# --- File: src/modules/meetings/core/session_plan/exporter/js_now_live.py ---
"""
Embedded JS (part 3/3): Live 'Now' detection, Now stream cards, Live Venue
Dispatch panel, dual-clock header pill and startup initialization.
"""

JS_NOW_LIVE = """\
  // ------------------------------------------------------------------
  // "Highlight Now" & Live Venue Dispatch Board ("Where should each delegate go?")
  // ------------------------------------------------------------------
  function getDistinctScheduledSlots() {
    const map = {};
    for (const ev of CAL_EVENTS) {
      const tus = (ev.sub_slot === 1) ? [1] : (ev.sub_slot === 2) ? [2] : [1, 2];
      for (const tu of tus) {
        const k = ev.day_idx + '|' + ev.quarter + '|' + tu;
        if (!map[k]) {
          map[k] = {
            key: k,
            day_idx: ev.day_idx,
            day_name: ev.day_name || DAY_NAMES_LIST[ev.day_idx] || ('Day ' + (ev.day_idx + 1)),
            quarter: ev.quarter,
            tu: tu,
            start_local: ev.start_local,
            end_local: ev.end_local
          };
        }
      }
    }
    const list = Object.values(map);
    list.sort((a, b) => (a.day_idx - b.day_idx) || (PERIOD_LIST.indexOf(a.quarter) - PERIOD_LIST.indexOf(b.quarter)) || (a.tu - b.tu));
    return list;
  }

  function hhmmToMinutes(hhmm) {
    const m = String(hhmm || '').match(/(\\d{1,2}):(\\d{2})/);
    if (!m) return 0;
    return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  }

  function detectLiveOrNextSlot() {
    const allSlots = getDistinctScheduledSlots();
    if (allSlots.length === 0) return null;

    if (!MEETING_BASE_DATE) {
      return Object.assign({}, allSlots[0], { is_live: false, status_note: 'Previewing first scheduled session' });
    }

    const now = new Date();
    const venueMs = now.getTime() + (MEETING_TZ_OFFSET_MIN * 60000);
    const vDate = new Date(venueMs);
    const pad = n => String(n).padStart(2, '0');
    const curTime = pad(vDate.getUTCHours()) + ':' + pad(vDate.getUTCMinutes());
    const curMin = vDate.getUTCHours() * 60 + vDate.getUTCMinutes();

    const parts = MEETING_BASE_DATE.split('-');
    const baseMs = Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    const vDayMs = Date.UTC(vDate.getUTCFullYear(), vDate.getUTCMonth(), vDate.getUTCDate());
    const dayDiff = Math.floor((vDayMs - baseMs) / 86400000);

    if (dayDiff >= 0 && dayDiff <= 4) {
      const todayEvents = CAL_EVENTS.filter(e => e.day_idx === dayDiff);
      const liveEvs = todayEvents.filter(e => e.start_local <= curTime && curTime < e.end_local);
      if (liveEvs.length > 0) {
        const sample = liveEvs.find(e => e.sub_slot === 1 || e.sub_slot === 2) || liveEvs[0];
        const tu = sample.sub_slot || 1;
        const rem = Math.max(0, hhmmToMinutes(sample.end_local) - curMin);
        return {
          key: dayDiff + '|' + sample.quarter + '|' + tu,
          day_idx: dayDiff,
          day_name: DAY_NAMES_LIST[dayDiff] || 'Today',
          quarter: sample.quarter,
          tu: tu,
          start_local: sample.start_local,
          end_local: sample.end_local,
          is_live: true,
          status_note: '🟢 LIVE NOW (Venue ' + curTime + ' · ' + rem + 'm left)'
        };
      }
      const nextEvs = todayEvents.filter(e => e.start_local > curTime).sort((a, b) => a.start_local.localeCompare(b.start_local));
      if (nextEvs.length > 0) {
        const n = nextEvs[0];
        const tu = n.sub_slot || 1;
        const wait = Math.max(0, hhmmToMinutes(n.start_local) - curMin);
        return {
          key: dayDiff + '|' + n.quarter + '|' + tu,
          day_idx: dayDiff,
          day_name: DAY_NAMES_LIST[dayDiff] || 'Today',
          quarter: n.quarter,
          tu: tu,
          start_local: n.start_local,
          end_local: n.end_local,
          is_live: false,
          status_note: '⏳ Up Next in ' + wait + 'm at ' + n.start_local + ' (Venue ' + curTime + ')'
        };
      }
      const tomorrowSlot = allSlots.find(s => s.day_idx > dayDiff);
      if (tomorrowSlot) {
        return Object.assign({}, tomorrowSlot, {
          is_live: false,
          status_note: '🌙 Today concluded — Previewing next day (' + tomorrowSlot.day_name + ')'
        });
      }
    }
    return Object.assign({}, allSlots[0], {
      is_live: false,
      status_note: (dayDiff < 0) ? '📅 Before meeting start — previewing Day 1' : '🏁 Outside live hours — schedule preview'
    });
  }

  function highlightNow(forceOpen) {
    if (forceOpen === false) {
      nowPanelOpen = false;
      const p = document.getElementById('nowDispatchPanel');
      if (p) p.style.display = 'none';
      document.querySelectorAll('.is-now-focused').forEach(el => el.classList.remove('is-now-focused'));
      if (currentView === 'now') switchView('grid');
      return;
    }
    nowPanelOpen = true;
    if (!nowTargetSlot) {
      nowTargetSlot = detectLiveOrNextSlot();
    }
    focusSlotCellsInDOM(nowTargetSlot, false);
    switchView('now');
  }

  function selectNowSlotByKey(slotKey) {
    const allSlots = getDistinctScheduledSlots();
    const found = allSlots.find(s => s.key === slotKey);
    if (found) {
      nowAutoFollow = false;
      nowTargetSlot = Object.assign({}, found, { is_live: false, status_note: '🔎 Manual Slot Inspection' });
      renderNowView();
      focusSlotCellsInDOM(nowTargetSlot, false);
    }
  }

  function stepNowSlot(delta) {
    const allSlots = getDistinctScheduledSlots();
    if (allSlots.length === 0) return;
    if (!nowTargetSlot) nowTargetSlot = detectLiveOrNextSlot();
    let idx = allSlots.findIndex(s => s.key === (nowTargetSlot && nowTargetSlot.key));
    if (idx < 0) idx = 0;
    idx = (idx + delta + allSlots.length) % allSlots.length;
    nowAutoFollow = false;
    nowTargetSlot = Object.assign({}, allSlots[idx], { is_live: false, status_note: '🔎 Manual Slot Inspection' });
    renderNowView();
    focusSlotCellsInDOM(nowTargetSlot, false);
  }

  function resetNowToLiveClock() {
    nowAutoFollow = true;
    nowTargetSlot = detectLiveOrNextSlot();
    renderNowView();
    focusSlotCellsInDOM(nowTargetSlot, false);
  }

  function focusSlotCellsInDOM(slotObj, shouldScroll) {
    document.querySelectorAll('.is-now-focused').forEach(el => el.classList.remove('is-now-focused'));
    if (!slotObj) return;
    const sDay = String(slotObj.day_idx);
    const sPeriod = String(slotObj.quarter);
    const gridCells = document.querySelectorAll('.grid-cell[data-day="' + sDay + '"][data-period="' + sPeriod + '"]');
    gridCells.forEach(el => {
      const cTu = el.getAttribute('data-tu');
      if (!cTu || cTu === '0' || cTu === String(slotObj.tu)) {
        el.classList.add('is-now-focused');
      }
    });
    const cardRows = document.querySelectorAll('.cards-slot-row[data-day="' + sDay + '"][data-period="' + sPeriod + '"]');
    cardRows.forEach(el => {
      const cTu = el.getAttribute('data-tu');
      if (!cTu || cTu === '0' || cTu === String(slotObj.tu)) {
        el.classList.add('is-now-focused');
      }
    });

    if (shouldScroll) {
      if (currentView === 'cards') {
        const dayName = DAY_NAMES_LIST[slotObj.day_idx] || 'Monday';
        switchDay(dayName);
        if (cardRows.length > 0 && cardRows[0].scrollIntoView) {
          cardRows[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      } else if (currentView === 'grid') {
        if (gridCells.length > 0 && gridCells[0].scrollIntoView) {
          gridCells[0].scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        }
      }
    }
  }

  function buildNowDelBadges(ev) {
    if (ev.is_all) {
      return '<span class="del-pill-badge is-all">👥 ALL</span>';
    }
    const dels = ev.delegates || [];
    if (dels.length > 0) {
      return dels.map(dn => {
        const meta = getRosterMemberMeta(dn);
        return '<span class="del-pill-badge" style="background:' + meta.bg_tint + ';color:' + meta.color + ';border:1px solid ' + meta.border_tint + ';">' + escHtml(meta.initials || dn.substring(0, 2).toUpperCase()) + '</span>';
      }).join(' ');
    }
    if (ev.not_followed) {
      return '<span class="del-pill-badge is-not-followed-pill" title="Out of scope">⛔</span>';
    }
    return '';
  }

  function renderNowView() {
    const container = document.getElementById('nowViewContent');
    if (!container) return;
    const allSlots = getDistinctScheduledSlots();
    if (!nowTargetSlot && allSlots.length > 0) {
      nowTargetSlot = detectLiveOrNextSlot();
    }
    if (!nowTargetSlot) {
      container.innerHTML = '<div class="empty-state-box">📭 No scheduled sessions available in this snapshot.</div>';
      return;
    }

    let optionsHtml = '';
    for (const s of allSlots) {
      const sel = (s.key === nowTargetSlot.key) ? ' selected' : '';
      optionsHtml += '<option value="' + escHtml(s.key) + '"' + sel + '>' +
        escHtml(s.day_name) + ' · ' + escHtml(s.quarter) + ' (TU ' + s.tu + ') — ' + escHtml(s.start_local) + '–' + escHtml(s.end_local) +
        '</option>';
    }

    const activeEvs = CAL_EVENTS.filter(e =>
      e.day_idx === nowTargetSlot.day_idx &&
      e.quarter === nowTargetSlot.quarter &&
      (e.sub_slot === 0 || e.sub_slot === nowTargetSlot.tu)
    ).sort((a, b) => a.stream_num - b.stream_num);

    const tu2ByStream = {};
    if (nowTargetSlot.tu === 1) {
      CAL_EVENTS.filter(e => e.day_idx === nowTargetSlot.day_idx && e.quarter === nowTargetSlot.quarter && e.sub_slot === 2)
        .forEach(e => { tu2ByStream[e.stream_num] = e; });
    }

    const headerBar = '<div class="aux-header-card" style="margin-bottom:10px;">' +
      '<div>' +
        '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">' +
          '<span style="font-size:11.5px;font-weight:700;padding:3px 10px;border-radius:999px;background:' + (nowTargetSlot.is_live ? '#DCFCE7;color:#15803D;border:1px solid #86EFAC;' : '#EFF6FF;color:#1D4ED8;border:1px solid #BFDBFE;') + '">' + escHtml(nowTargetSlot.status_note || 'Schedule Preview') + '</span>' +
          '<h2 class="aux-title" style="margin:0;">' + escHtml(nowTargetSlot.day_name) + ' · ' + escHtml(nowTargetSlot.quarter) + ' (TU ' + nowTargetSlot.tu + ')</h2>' +
          '<span style="font-size:12px;color:#475569;font-weight:600;">🕒 ' + escHtml(nowTargetSlot.start_local || '') + '–' + escHtml(nowTargetSlot.end_local || '') + ' Venue</span>' +
        '</div>' +
      '</div>' +
      '<div class="now-slot-controls">' +
        '<button class="now-step-btn" onclick="stepNowSlot(-1)" title="Previous Time Unit">◀ Prev</button>' +
        '<select class="now-select" onchange="selectNowSlotByKey(this.value)">' + optionsHtml + '</select>' +
        '<button class="now-step-btn" onclick="stepNowSlot(1)" title="Next Time Unit">Next ▶</button>' +
        '<button class="now-step-btn" onclick="resetNowToLiveClock()" title="Sync to Live Venue Clock" style="background:#ECFDF5;color:#047857;border-color:#6EE7B7;">🔴 Live Now</button>' +
      '</div>' +
    '</div>';

    // Section 1: Happening Now — Rooms & Active Streams
    let sec1Html = '<div class="now-section-title">' + (nowTargetSlot.is_live ? '🔴 Happening Now — Rooms &amp; Active Streams' : '📍 Active Window — Rooms &amp; Streams') + '</div>';
    const isBreakOnly = activeEvs.length > 0 && activeEvs.every(e => e.is_break || e.stream_num === 0);
    if (isBreakOnly) {
      sec1Html += '<div class="now-streams-row" style="grid-template-columns:1fr;">';
      for (const ev of activeEvs) {
        sec1Html += renderNowStreamCardHtml(ev.stream_num, ev, null);
      }
      sec1Html += '</div>';
    } else {
      sec1Html += '<div class="now-streams-row">';
      for (const sNum of [1, 2, 3]) {
        const ev = activeEvs.find(e => e.stream_num === sNum) || null;
        const tu2Ev = tu2ByStream[sNum] || null;
        sec1Html += renderNowStreamCardHtml(sNum, ev, tu2Ev);
      }
      sec1Html += '</div>';
    }

    // Section 2 & 3: Later Periods Today (Up Next 1-2 + Rest of Today)
    const curPIdx = PERIOD_LIST.indexOf(nowTargetSlot.quarter);
    const laterPeriods = [];
    for (let i = curPIdx + 1; i < PERIOD_LIST.length; i++) {
      const pName = PERIOD_LIST[i];
      const pEvs = CAL_EVENTS.filter(e => e.day_idx === nowTargetSlot.day_idx && e.quarter === pName)
        .sort((a, b) => (a.stream_num - b.stream_num) || (a.sub_slot - b.sub_slot));
      if (pEvs.length > 0) {
        const sTimes = pEvs.map(e => e.start_local).filter(Boolean).sort();
        const eTimes = pEvs.map(e => e.end_local).filter(Boolean).sort();
        laterPeriods.push({
          period: pName,
          start_local: sTimes[0] || '',
          end_local: eTimes[eTimes.length - 1] || '',
          is_break: pEvs.every(e => e.is_break),
          events: pEvs
        });
      }
    }

    const upNext = laterPeriods.slice(0, 2);
    const restDay = laterPeriods.slice(2);

    let sec2Html = '<div class="now-section-title">⏭️ Up Next Today (' + escHtml(nowTargetSlot.day_name) + ')</div>';
    if (upNext.length === 0) {
      sec2Html += '<div class="empty-state-box" style="padding:14px;">✅ No further sessions scheduled later on ' + escHtml(nowTargetSlot.day_name) + '.</div>';
    } else {
      for (const lp of upNext) {
        const pIcon = lp.is_break ? '☕' : '⏭️';
        let subCards = '';
        for (const ev of lp.events) {
          const tuTag = (ev.sub_slot === 1) ? ' [TU 1]' : (ev.sub_slot === 2) ? ' [TU 2]' : '';
          const locPill = ev.location_note ? (' <span class="now-loc-badge" style="font-size:10px;padding:1px 5px;">🧭 ' + escHtml(ev.location_note) + '</span>') : '';
          const swTag = ev.room_swapped ? ' <span class="now-swap-badge">⇄</span>' : '';
          subCards += '<div class="now-upcoming-sub" style="cursor:pointer;" onclick="jumpToSlotInGrid(' + ev.day_idx + ', \\'' + escHtml(ev.quarter) + '\\', ' + ev.stream_num + ')">' +
            '<div style="font-size:11px;font-weight:700;color:#1E40AF;margin-bottom:2px;">Str ' + ev.stream_num + tuTag + ' · 📍 ' + escHtml(ev.room) + swTag + ' ' + locPill + '</div>' +
            '<div style="font-size:11.5px;font-weight:500;color:#0F172A;">' + escHtml(ev.title) + '</div>' +
          '</div>';
        }
        sec2Html += '<div class="now-upcoming-card">' +
          '<div class="now-upcoming-hdr"><span>' + pIcon + ' ' + escHtml(lp.period) + ' · 🕒 ' + escHtml(lp.start_local) + '–' + escHtml(lp.end_local) + '</span></div>' +
          '<div class="now-upcoming-grid">' + subCards + '</div>' +
        '</div>';
      }
    }

    let sec3Html = '';
    if (restDay.length > 0) {
      sec3Html += '<div class="now-section-title">📅 Rest of ' + escHtml(nowTargetSlot.day_name) + ' Timeline</div>';
      for (const rp of restDay) {
        const lineParts = rp.events.map(ev => {
          const locS = ev.location_note ? (' (' + escHtml(ev.location_note) + ')') : '';
          const tuS = (ev.sub_slot === 1 || ev.sub_slot === 2) ? (' TU' + ev.sub_slot) : '';
          return '<strong>Str ' + ev.stream_num + tuS + ' [📍 ' + escHtml(ev.room) + locS + ']:</strong> ' + escHtml((ev.title || '').substring(0, 75));
        }).join(' &nbsp;|&nbsp; ');
        const fEv = rp.events[0];
        sec3Html += '<div class="now-timeline-row">' +
          '<div style="min-width:140px;font-size:11.5px;"><strong>' + escHtml(rp.period) + '</strong><br><span style="color:#64748B;font-size:10.5px;">' + escHtml(rp.start_local) + '–' + escHtml(rp.end_local) + '</span></div>' +
          '<div style="flex:1;font-size:11.5px;color:#334155;">' + lineParts + '</div>' +
          (fEv ? '<button class="btn-jump-grid" onclick="jumpToSlotInGrid(' + fEv.day_idx + ', \\'' + escHtml(fEv.quarter) + '\\', ' + fEv.stream_num + ')">🔍 Grid</button>' : '') +
        '</div>';
      }
    }

    container.innerHTML = headerBar + sec1Html + sec2Html + sec3Html;
    renderNowDispatchPanel();
  }

  function renderNowStreamCardHtml(sNum, ev, tu2Ev) {
    const sMeta = (typeof STREAMS_META !== 'undefined' && STREAMS_META[String(sNum)]) ? STREAMS_META[String(sNum)] : null;
    const strLabel = sNum > 0 ? ('🏢 Stream ' + sNum) : '🏢 Plenary / Break';
    if (!ev) {
      const defRoom = sMeta ? sMeta.room_name : ('Stream ' + sNum);
      const defLoc = sMeta && sMeta.location_note ? (' · 🧭 ' + escHtml(sMeta.location_note)) : '';
      return '<div class="now-stream-card is-idle">' +
        '<div class="now-card-head"><span class="stream-badge str-' + sNum + '">' + strLabel + '</span></div>' +
        '<div style="font-size:11.5px;color:#64748B;">📍 ' + escHtml(defRoom) + defLoc + '</div>' +
        '<div style="color:#94A3B8;font-style:italic;font-size:12px;padding:12px 0;">— No active session scheduled in this stream —</div>' +
      '</div>';
    }

    const delHtml = buildNowDelBadges(ev);
    const swapBadge = ev.room_swapped ? '<span class="now-swap-badge">⇄ Moved Room</span>' : '';
    const locBadge = ev.location_note ? '<span class="now-loc-badge">🧭 ' + escHtml(ev.location_note) + '</span>' : '';
    const metaParts = [];
    if (ev.convener) metaParts.push('👤 ' + escHtml(ev.convener));
    if (ev.tohru) metaParts.push('📡 ' + escHtml(ev.tohru));

    let tu2Html = '';
    if (tu2Ev && tu2Ev.slot_key !== ev.slot_key) {
      const tu2Loc = tu2Ev.location_note ? (' (' + escHtml(tu2Ev.location_note) + ')') : '';
      tu2Html = '<div class="now-tu2-hint">⏭️ <strong>2nd Half (' + escHtml(tu2Ev.start_local) + '–' + escHtml(tu2Ev.end_local) + '):</strong> 📍 ' + escHtml(tu2Ev.room) + tu2Loc + ' — ' + escHtml((tu2Ev.title || '').substring(0, 80)) + '</div>';
    }

    const noteHtml = ev.note ? ('<div class="note-chip">📝 ' + escHtml(ev.note) + '</div>') : '';

    return '<div class="now-stream-card str-' + sNum + '">' +
      '<div class="now-card-head">' +
        '<span class="stream-badge str-' + sNum + '">' + strLabel + (ev.is_drafting ? ' · Drafting' : '') + '</span>' +
        '<div style="display:flex;gap:3px;flex-wrap:wrap;">' + delHtml + '</div>' +
      '</div>' +
      '<div class="now-room-box">' +
        '<div class="now-room-title"><span>📍 ' + escHtml(ev.room) + '</span>' + swapBadge + '</div>' +
        (locBadge ? '<div>' + locBadge + '</div>' : '') +
        (metaParts.length ? '<div class="now-room-meta">' + metaParts.join(' · ') + '</div>' : '') +
      '</div>' +
      '<div class="now-topic-text">' + escHtml(ev.title) + '</div>' +
      noteHtml +
      tu2Html +
      '<div style="display:flex;justify-content:flex-end;margin-top:4px;">' +
        '<button class="btn-jump-grid" onclick="jumpToSlotInGrid(' + ev.day_idx + ', \\'' + escHtml(ev.quarter) + '\\', ' + ev.stream_num + ')">🔍 Show in Grid</button>' +
      '</div>' +
    '</div>';
  }

  function renderNowDispatchPanel() {
    const panel = document.getElementById('nowDispatchPanel');
    if (!panel) return;
    const allSlots = getDistinctScheduledSlots();
    if (!nowTargetSlot && allSlots.length > 0) {
      nowTargetSlot = detectLiveOrNextSlot();
    }
    if (!nowTargetSlot) {
      panel.style.display = 'none';
      return;
    }

    const activeEvs = CAL_EVENTS.filter(e =>
      e.day_idx === nowTargetSlot.day_idx &&
      e.quarter === nowTargetSlot.quarter &&
      (e.sub_slot === 0 || e.sub_slot === nowTargetSlot.tu)
    ).sort((a, b) => a.stream_num - b.stream_num);

    const allMembers = getAllRosterNames();
    let cardsHtml = '';

    for (const member of allMembers) {
      const mMeta = getRosterMemberMeta(member);
      const mLower = member.toLowerCase();
      const myEvs = activeEvs.filter(e => e.is_all || (e.delegates || []).some(d => d.toLowerCase() === mLower));

      if (myEvs.length === 0) {
        cardsHtml += '<div class="now-del-card status-free" onclick="selectDelegate(\\'' + escHtml(member) + '\\')">' +
          '<div class="now-del-top">' +
            '<span class="now-del-name" style="color:' + mMeta.color + ';">● ' + escHtml(member) + '</span>' +
            '<span class="now-del-status-tag" style="background:#DCFCE7;color:#15803D;">🟢 Free</span>' +
          '</div>' +
          '<div class="now-del-dest" style="color:#64748B;">No assigned stream</div>' +
        '</div>';
      } else if (myEvs.length === 1) {
        const ev = myEvs[0];
        const roleBadge = ev.is_all ? '👥 All Team' : (ev.is_cov ? '⚡ Override' : '📌 Owner');
        const locS = ev.location_note ? (' (' + ev.location_note + ')') : '';
        cardsHtml += '<div class="now-del-card status-busy" onclick="jumpToSlotInGrid(' + ev.day_idx + ', \\'' + escHtml(ev.quarter) + '\\', ' + ev.stream_num + ')">' +
          '<div class="now-del-top">' +
            '<span class="now-del-name" style="color:' + mMeta.color + ';">● ' + escHtml(member) + '</span>' +
            '<span class="now-del-status-tag" style="background:#DBEAFE;color:#1D4ED8;">➡️ Str ' + ev.stream_num + ' (' + roleBadge + ')</span>' +
          '</div>' +
          '<div class="now-del-dest">📍 ' + escHtml(ev.room + locS) + '</div>' +
          '<div class="now-del-topic" title="' + escHtml(ev.title) + '">' + escHtml(ev.title) + '</div>' +
        '</div>';
      } else {
        const rooms = myEvs.map(e => 'Str ' + e.stream_num + ' (' + e.room + ')').join(' ⚡ ');
        const titles = myEvs.map(e => 'Str ' + e.stream_num + ': ' + e.title).join(' | ');
        cardsHtml += '<div class="now-del-card status-clash" onclick="switchView(\\'conflicts\\')">' +
          '<div class="now-del-top">' +
            '<span class="now-del-name" style="color:' + mMeta.color + ';">● ' + escHtml(member) + '</span>' +
            '<span class="now-del-status-tag" style="background:#FEE2E2;color:#B91C1C;">⚠️ CLASH (' + myEvs.length + ')</span>' +
          '</div>' +
          '<div class="now-del-dest" style="color:#B91C1C;">📍 ' + escHtml(rooms) + '</div>' +
          '<div class="now-del-topic" title="' + escHtml(titles) + '">' + escHtml(titles) + '</div>' +
        '</div>';
      }
    }

    if (allMembers.length === 0) {
      panel.style.display = 'none';
      return;
    }
    panel.style.display = 'block';
    panel.innerHTML =
      '<div class="now-panel-header" style="margin-bottom:6px;padding-bottom:5px;">' +
        '<div class="now-panel-title"><span>👥 Where Each Delegate Should Go</span></div>' +
      '</div>' +
      '<div class="now-dispatch-grid">' + cardsHtml + '</div>';
  }

  // Live Venue Status Monitor ("Now / Up Next") & Header Dual Clock
  function updateHeaderClockPill(vDate, nowDate) {
    const pill = document.getElementById('liveClockPill');
    if (!pill) return;
    const pad = n => String(n).padStart(2, '0');
    const dayShort = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const vStr = dayShort[vDate.getUTCDay()] + ' ' + pad(vDate.getUTCHours()) + ':' + pad(vDate.getUTCMinutes());
    const lStr = pad(nowDate.getHours()) + ':' + pad(nowDate.getMinutes());
    pill.innerHTML = '🏛️ Venue: <strong>' + vStr + '</strong> <span style="color:#93C5FD;">|</span> 🏠 Local: <strong>' + lStr + '</strong>';
  }

  function checkLiveVenueStatus() {
    const banner = document.getElementById('liveVenueBanner');
    const nowTabBtn = document.getElementById('btnViewNow');
    const now = new Date();
    const venueMs = now.getTime() + (MEETING_TZ_OFFSET_MIN * 60000);
    const vDate = new Date(venueMs);
    updateHeaderClockPill(vDate, now);

    if (!MEETING_BASE_DATE) return;

    document.querySelectorAll('.is-live-now').forEach(el => el.classList.remove('is-live-now'));
    document.querySelectorAll('.is-live-row-hdr').forEach(el => el.classList.remove('is-live-row-hdr'));

    const pad = n => String(n).padStart(2, '0');
    const curTime = pad(vDate.getUTCHours()) + ':' + pad(vDate.getUTCMinutes());

    const parts = MEETING_BASE_DATE.split('-');
    const baseMs = Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    const vDayMs = Date.UTC(vDate.getUTCFullYear(), vDate.getUTCMonth(), vDate.getUTCDate());
    const dayDiff = Math.floor((vDayMs - baseMs) / 86400000);
    const closeBtnHtml = '<button type="button" class="live-banner-close" onclick="dismissLiveBanner(event)" title="Hide live banner">✕</button>';

    if (dayDiff >= 0 && dayDiff <= 4) {
      const todayEvents = CAL_EVENTS.filter(e => e.day_idx === dayDiff);
      const liveNow = todayEvents.filter(e => e.start_local <= curTime && curTime < e.end_local);
      const upcoming = todayEvents.filter(e => e.start_local > curTime).sort((a, b) => a.start_local.localeCompare(b.start_local));

      if (liveNow.length > 0) {
        const livePeriods = {};
        liveNow.forEach(ev => {
          const sDay = String(ev.day_idx);
          const sPeriod = String(ev.quarter);
          const sStr = String(ev.stream_num);
          livePeriods[sPeriod] = true;
          document.querySelectorAll('.grid-cell[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('is-live-now'));
          document.querySelectorAll('.cards-slot-row[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('is-live-now'));
        });
        Object.keys(livePeriods).forEach(lp => {
          document.querySelectorAll('tr.grid-row[data-period="' + lp + '"]').forEach(el => el.classList.add('is-live-row-hdr'));
        });
        if (nowTabBtn) nowTabBtn.textContent = '🟢 Now (' + liveNow[0].quarter + ')';
        if (banner && !liveBannerDismissed) {
          const liveTitles = liveNow.map(e => (e.is_drafting ? '🛠️ Draft: ' : ('Str ' + e.stream_num + ' [' + e.room + ']: ')) + (e.title || '').substring(0, 36)).join(' | ');
          const nextMsg = upcoming.length > 0 ? (' &bull; Next: ' + upcoming[0].start_local + ' (' + (upcoming[0].title || '').substring(0, 28) + ')') : '';
          banner.innerHTML = '<span><span class="pulse-dot"></span> <strong>LIVE NOW</strong> (' + curTime + '): ' + escHtml(liveTitles) + nextMsg + '</span><span class="live-banner-right"><span style="font-size:10.5px;text-decoration:underline;">Open Now Tab</span>' + closeBtnHtml + '</span>';
          banner.style.display = 'flex';
        }
      } else {
        if (nowTabBtn) nowTabBtn.textContent = '🔴 Now';
        if (banner && !liveBannerDismissed) {
          if (upcoming.length > 0) {
            const n = upcoming[0];
            banner.innerHTML = '<span>⏳ <strong>Venue ' + curTime + '</strong> &bull; Next at ' + n.start_local + ' in 📍 ' + escHtml(n.room) + ': ' + escHtml((n.title || '').substring(0, 42)) + '</span><span class="live-banner-right"><span style="font-size:10.5px;text-decoration:underline;">Open Now Tab</span>' + closeBtnHtml + '</span>';
          } else {
            banner.innerHTML = '<span>🌙 <strong>Venue ' + curTime + '</strong> &bull; Sessions concluded for today.</span><span class="live-banner-right"><span style="font-size:10.5px;text-decoration:underline;">Open Now Tab</span>' + closeBtnHtml + '</span>';
          }
          banner.style.display = 'flex';
        }
      }
    } else if (banner) {
      if (nowTabBtn) nowTabBtn.textContent = '🔴 Now';
      if (!liveBannerDismissed) {
        if (dayDiff < 0) {
          const daysUntil = -dayDiff;
          banner.innerHTML = '<span>📅 Meeting starts in ' + daysUntil + ' day' + (daysUntil > 1 ? 's' : '') + ' (' + MEETING_BASE_DATE + ')</span><span class="live-banner-right"><span style="font-size:10.5px;text-decoration:underline;">Preview Now Tab</span>' + closeBtnHtml + '</span>';
        } else {
          banner.innerHTML = '<span>🏁 Meeting concluded (' + MEETING_BASE_DATE + ')</span><span class="live-banner-right"><span style="font-size:10.5px;text-decoration:underline;">Preview Now Tab</span>' + closeBtnHtml + '</span>';
        }
        banner.style.display = 'flex';
      }
    }

    if (nowAutoFollow && currentView === 'now') {
      nowTargetSlot = detectLiveOrNextSlot();
      renderNowView();
    }
  }

  // Auto-switch to mobile cards on narrow screens
  if (window.innerWidth <= 768) {
    switchView('cards');
  }
  if (window.addEventListener) {
    window.addEventListener('resize', updateStickyHeaderOffsets);
  }
  detectDelegateClashes('ALL');
  checkLiveVenueStatus();
  updateStickyHeaderOffsets();
  setInterval(checkLiveVenueStatus, 30000);
"""
