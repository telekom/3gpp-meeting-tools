# --- File: src/modules/meetings/core/session_plan/exporter/js_core.py ---
"""
Embedded JS (part 1/3): UI state, HTML escaping, view switcher, day/delegate/
stream/search filters, slot inspector drawer and client-side .ics export.
"""

JS_CORE = """
  let currentView = 'grid';
  let selectedDelegate = 'ALL';
  let onlyUnassigned = false;
  let onlyDrafting = false;
  let unassignedScheduledOnly = true;
  let nowAutoFollow = true;
  let nowPanelOpen = false;
  let nowTargetSlot = null; // { day_idx, quarter, tu, label, is_live }
  let dayFilterMode = 'ALL'; // 'ALL', 'TODAY_PLUS', or 'CUSTOM'
  let selectedDaysSet = new Set();
  let liveBannerDismissed = false;

  const PERIOD_LIST = ['Q0', 'Break 1', 'Q1', 'Coffee 1', 'Q2', 'Lunch', 'Q3', 'Coffee 2', 'Q4', 'Q5'];
  const DAY_NAMES_LIST = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

  function escHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function updateStickyHeaderOffsets() {
    const dayRow = document.querySelector('.hdr-days-row th:not(.corner-hdr)');
    if (dayRow && dayRow.offsetHeight && document.documentElement && document.documentElement.style && document.documentElement.style.setProperty) {
      document.documentElement.style.setProperty('--day-hdr-h', dayRow.offsetHeight + 'px');
    }
  }

  function getLiveVenueDayIndex() {
    if (!MEETING_BASE_DATE) return -1;
    const now = new Date();
    const venueMs = now.getTime() + (MEETING_TZ_OFFSET_MIN * 60000);
    const vDate = new Date(venueMs);
    const parts = MEETING_BASE_DATE.split('-');
    const baseMs = Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    const vDayMs = Date.UTC(vDate.getUTCFullYear(), vDate.getUTCMonth(), vDate.getUTCDate());
    const diff = Math.floor((vDayMs - baseMs) / 86400000);
    return (diff >= 0 && diff <= 4) ? diff : -1;
  }

  function selectDayFilter(target) {
    if (target === 'ALL') {
      dayFilterMode = 'ALL';
      selectedDaysSet.clear();
    } else if (target === 'TODAY_PLUS') {
      dayFilterMode = 'TODAY_PLUS';
      selectedDaysSet.clear();
    } else {
      const dIdx = parseInt(target, 10);
      if (dayFilterMode !== 'CUSTOM') {
        dayFilterMode = 'CUSTOM';
        selectedDaysSet = new Set([dIdx]);
      } else {
        if (selectedDaysSet.has(dIdx)) {
          selectedDaysSet.delete(dIdx);
        } else {
          selectedDaysSet.add(dIdx);
        }
        if (selectedDaysSet.size === 0 || selectedDaysSet.size >= 5) {
          dayFilterMode = 'ALL';
          selectedDaysSet.clear();
        }
      }
    }

    let visibleDays = new Set([0, 1, 2, 3, 4]);
    if (dayFilterMode === 'TODAY_PLUS') {
      const todayIdx = getLiveVenueDayIndex();
      if (todayIdx >= 0) {
        visibleDays = new Set();
        for (let i = todayIdx; i < 5; i++) visibleDays.add(i);
      }
    } else if (dayFilterMode === 'CUSTOM' && selectedDaysSet.size > 0) {
      visibleDays = new Set(selectedDaysSet);
    }

    document.querySelectorAll('.day-filter-btn').forEach(btn => {
      const val = btn.getAttribute('data-day-filter');
      if (val === 'ALL') {
        btn.classList.toggle('active', dayFilterMode === 'ALL');
      } else if (val === 'TODAY_PLUS') {
        btn.classList.toggle('active', dayFilterMode === 'TODAY_PLUS');
      } else {
        btn.classList.toggle('active', dayFilterMode === 'CUSTOM' && selectedDaysSet.has(parseInt(val, 10)));
      }
    });

    const ov = document.getElementById('overviewView');
    if (ov) {
      ov.querySelectorAll('[data-day]').forEach(el => {
        const d = parseInt(el.getAttribute('data-day'), 10);
        el.classList.toggle('day-hidden', !visibleDays.has(d));
      });
      const tbl = ov.querySelector('.overview-table');
      if (tbl) {
        tbl.classList.toggle('filtered-days', visibleDays.size < 5);
        if (tbl.style) {
          if (visibleDays.size < 5) {
            const isMobile = Boolean(window.innerWidth && window.innerWidth <= 768);
            tbl.style.minWidth = ((isMobile ? 108 : 152) + visibleDays.size * (isMobile ? 236 : 180)) + 'px';
          } else {
            tbl.style.minWidth = '';
          }
        }
      }
    }
    updateStickyHeaderOffsets();
  }

  function dismissLiveBanner(ev) {
    if (ev && ev.stopPropagation) ev.stopPropagation();
    liveBannerDismissed = true;
    const banner = document.getElementById('liveVenueBanner');
    if (banner) banner.style.display = 'none';
  }

  function switchView(viewName) {
    currentView = viewName;
    const views = ['grid', 'now', 'cards', 'conflicts', 'unassigned'];
    const btnIds = {
      grid: 'btnViewGrid',
      now: 'btnViewNow',
      cards: 'btnViewCards',
      conflicts: 'btnViewConflicts',
      unassigned: 'btnViewUnassigned'
    };
    const divIds = {
      grid: 'overviewView',
      now: 'nowView',
      cards: 'cardsView',
      conflicts: 'conflictsView',
      unassigned: 'unassignedView'
    };

    for (const v of views) {
      const btn = document.getElementById(btnIds[v]);
      if (btn) btn.classList.toggle('active', v === viewName);
      const div = document.getElementById(divIds[v]);
      if (div) div.style.display = (v === viewName) ? (v === 'grid' ? 'flex' : 'block') : 'none';
    }

    const gc = document.getElementById('gridControls');
    if (gc) gc.style.display = (viewName === 'grid') ? 'flex' : 'none';
    const dc = document.getElementById('dayFilterStrip');
    if (dc) dc.style.display = (viewName === 'grid') ? 'inline-flex' : 'none';

    if (viewName === 'now') {
      renderNowView();
    } else if (viewName === 'conflicts') {
      renderConflictsView();
    } else if (viewName === 'unassigned') {
      renderUnassignedView();
    } else {
      applyFilters();
      if (viewName === 'grid') updateStickyHeaderOffsets();
    }
  }

  function selectDelegate(delName) {
    selectedDelegate = delName;
    document.querySelectorAll('.del-filter-btn[data-del]').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-del') === delName);
    });
    const dlBtn = document.getElementById('btnDownloadIcs');
    if (dlBtn) {
      dlBtn.textContent = (delName === 'ALL') ? '📅 Team .ics' : ('📅 ' + delName + ' .ics');
    }
    applyFilters();
    detectDelegateClashes(delName);
    if (currentView === 'conflicts') renderConflictsView();
    if (nowPanelOpen) renderNowDispatchPanel();
  }

  function toggleUnassigned() {
    onlyUnassigned = !onlyUnassigned;
    const btn = document.getElementById('btnUnassigned');
    if (btn) btn.classList.toggle('active', onlyUnassigned);
    if (document.body && document.body.classList) {
      document.body.classList.toggle('filter-unassigned', onlyUnassigned);
    }
    applyFilters();
  }

  function toggleDrafting() {
    onlyDrafting = !onlyDrafting;
    const btn = document.getElementById('btnDrafting');
    if (btn) btn.classList.toggle('active', onlyDrafting);
    applyFilters();
  }

  function toggleStream(num) {
    const chk = document.getElementById('chkStr' + num);
    const rows = document.querySelectorAll('.stream-row-' + num);
    rows.forEach(r => r.style.display = (chk && chk.checked) ? '' : 'none');
  }

  function toggleBreaks() {
    const chk = document.getElementById('chkBreaks');
    const rows = document.querySelectorAll('.break-row');
    rows.forEach(r => r.style.display = (chk && chk.checked) ? '' : 'none');
  }

  function switchDay(dayName, btn) {
    document.querySelectorAll('.day-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.day-section').forEach(s => s.style.display = 'none');
    const activeBtn = btn || (typeof event !== 'undefined' && event && event.target) || document.querySelector('.day-tab[data-day="' + dayName + '"]');
    if (activeBtn) activeBtn.classList.add('active');
    const target = document.getElementById('day-' + dayName);
    if (target) target.style.display = 'block';
  }

  function jumpToSlotInGrid(dayIdx, period, streamNum) {
    const dayName = DAY_NAMES_LIST[dayIdx] || 'Monday';
    if (window.innerWidth && window.innerWidth <= 768) {
      switchView('cards');
      switchDay(dayName);
      const sel = '.cards-slot-row[data-day="' + dayIdx + '"][data-period="' + period + '"]' + (streamNum ? '[data-stream="' + streamNum + '"]' : '');
      const el = document.querySelector(sel);
      if (el && el.scrollIntoView) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else {
      switchView('grid');
      const sel = '.grid-cell[data-day="' + dayIdx + '"][data-period="' + period + '"]' + (streamNum ? '[data-stream="' + streamNum + '"]' : '');
      const el = document.querySelector(sel);
      if (el) {
        if (el.classList && el.classList.contains('day-hidden')) {
          selectDayFilter('ALL');
        }
        if (el.scrollIntoView) el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        el.classList.add('is-now-focused');
        setTimeout(() => el.classList.remove('is-now-focused'), 2600);
      }
    }
  }

  function applyFilters() {
    const sInput = document.getElementById('searchInput');
    const query = ((sInput && sInput.value) || '').toLowerCase().trim();

    if (currentView === 'conflicts') {
      renderConflictsView();
      return;
    }
    if (currentView === 'unassigned') {
      renderUnassignedView();
      return;
    }

    // 1. Filter Overview Grid Cells
    const gridCells = document.querySelectorAll('.grid-cell.slot-cell');
    gridCells.forEach(cell => {
      const sTxt = cell.getAttribute('data-search') || '';
      const isUnas = cell.getAttribute('data-unassigned') === '1';
      const isDrf = cell.getAttribute('data-drafting') === '1';
      const dels = (cell.getAttribute('data-delegates') || '').split(',').map(d => d.trim().toLowerCase()).filter(Boolean);
      const isAll = cell.getAttribute('data-is-all') === '1';

      const matchesSearch = !query || sTxt.includes(query);
      const matchesUnassigned = !onlyUnassigned || isUnas;
      const matchesDrafting = !onlyDrafting || isDrf;
      const matchesDelegate = onlyUnassigned ? true : ((selectedDelegate === 'ALL') || isAll || dels.includes(selectedDelegate.toLowerCase()));

      const isMatch = matchesSearch && matchesUnassigned && matchesDrafting && matchesDelegate;

      if (selectedDelegate === 'ALL' && !query && !onlyUnassigned && !onlyDrafting) {
        cell.classList.remove('dimmed', 'highlighted');
      } else if (isMatch) {
        cell.classList.remove('dimmed');
        cell.classList.add('highlighted');
      } else {
        cell.classList.add('dimmed');
        cell.classList.remove('highlighted');
      }
    });

    // 2. Filter Cards View Rows
    const cardRows = document.querySelectorAll('.cards-slot-row');
    cardRows.forEach(row => {
      const sTxt = row.getAttribute('data-search') || '';
      const isUnas = row.getAttribute('data-unassigned') === '1';
      const isDrf = row.getAttribute('data-drafting') === '1';
      const dels = (row.getAttribute('data-delegates') || '').split(',').map(d => d.trim().toLowerCase()).filter(Boolean);
      const isAll = row.getAttribute('data-is-all') === '1';

      const matchesSearch = !query || sTxt.includes(query);
      const matchesUnassigned = !onlyUnassigned || isUnas;
      const matchesDrafting = !onlyDrafting || isDrf;
      const matchesDelegate = onlyUnassigned ? true : ((selectedDelegate === 'ALL') || isAll || dels.includes(selectedDelegate.toLowerCase()));

      if (matchesSearch && matchesUnassigned && matchesDrafting && matchesDelegate) {
        row.classList.remove('hidden');
      } else {
        row.classList.add('hidden');
      }
    });

    // Hide empty quarter cards
    document.querySelectorAll('.quarter-card').forEach(card => {
      const visible = card.querySelectorAll('.cards-slot-row:not(.hidden)');
      card.style.display = (visible.length > 0) ? 'block' : 'none';
    });
  }

  // Click-Activated Slide-In Slot Inspector Drawer
  const inspectorDrawer = document.getElementById('slotInspectorDrawer');
  const inspectorBody = document.getElementById('slotInspectorBody');
  let inspectedCellEl = null;

  function openSlotInspector(cellEl, htmlContent) {
    if (!inspectorDrawer || !inspectorBody) return;
    if (inspectedCellEl && inspectedCellEl !== cellEl && inspectedCellEl.classList) {
      inspectedCellEl.classList.remove('is-inspected');
    }
    inspectedCellEl = cellEl || null;
    if (inspectedCellEl && inspectedCellEl.classList) {
      inspectedCellEl.classList.add('is-inspected');
    }
    inspectorBody.innerHTML = htmlContent || '';
    inspectorBody.scrollTop = 0;
    inspectorDrawer.classList.add('open');
    if (inspectorDrawer.setAttribute) inspectorDrawer.setAttribute('aria-hidden', 'false');
  }

  function closeSlotInspector() {
    if (inspectedCellEl && inspectedCellEl.classList) {
      inspectedCellEl.classList.remove('is-inspected');
    }
    inspectedCellEl = null;
    if (inspectorDrawer) {
      inspectorDrawer.classList.remove('open');
      if (inspectorDrawer.setAttribute) inspectorDrawer.setAttribute('aria-hidden', 'true');
    }
  }

  document.querySelectorAll('.slot-cell, .cards-slot-row').forEach(cell => {
    const dataEl = cell.querySelector('.slot-tooltip-data');
    if (!dataEl) return;

    cell.addEventListener('click', e => {
      if (e && e.target && e.target.closest && e.target.closest('a, button')) return;
      if (inspectedCellEl === cell && inspectorDrawer && inspectorDrawer.classList.contains('open')) {
        closeSlotInspector();
        return;
      }
      openSlotInspector(cell, dataEl.innerHTML);
    });
  });

  if (document && document.addEventListener) {
    document.addEventListener('keydown', e => {
      if (e && e.key === 'Escape' && inspectorDrawer && inspectorDrawer.classList.contains('open')) {
        closeSlotInspector();
      }
    });

    document.addEventListener('click', e => {
      if (!inspectorDrawer || !inspectorDrawer.classList.contains('open')) return;
      if (e && e.target && e.target.closest && e.target.closest('#slotInspectorDrawer, .slot-cell, .cards-slot-row')) {
        return;
      }
      closeSlotInspector();
    });
  }

  // RFC 5545 Client-Side .ics Calendar Generation & Download
  function venueToUtcIso(baseDateStr, dayIdx, timeStr, offsetMin) {
    if (!baseDateStr) return null;
    const parts = baseDateStr.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10) + dayIdx;
    const tParts = (timeStr || '09:00').split(':');
    const h = parseInt(tParts[0], 10);
    const m = parseInt(tParts[1], 10);

    const utcMs = Date.UTC(year, month, day, h, m) - (offsetMin * 60000);
    const d = new Date(utcMs);
    const pad = n => String(n).padStart(2, '0');
    return d.getUTCFullYear() +
           pad(d.getUTCMonth() + 1) +
           pad(d.getUTCDate()) + 'T' +
           pad(d.getUTCHours()) +
           pad(d.getUTCMinutes()) + '00Z';
  }

  function downloadSelectedCalendar() {
    if (!MEETING_BASE_DATE) {
      alert('Meeting start date is not configured for this snapshot; cannot calculate physical calendar dates.');
      return;
    }
    const isAll = (selectedDelegate === 'ALL');
    const filterName = isAll ? 'Whole Team' : selectedDelegate;
    const safeName = isAll ? 'Team' : selectedDelegate.replace(/[^a-zA-Z0-9_-]/g, '_');
    const nowUtc = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';

    const calLines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//3GPP Meeting Tools//SA2 #' + MEETING_NUMBER + ' Session Plan//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'X-WR-CALNAME:SA2 #' + MEETING_NUMBER + ' - ' + filterName + ' (' + TDOC_ID + ')'
    ];

    let count = 0;
    for (const ev of CAL_EVENTS) {
      const isSlotAll = ev.is_all;
      const effDels = ev.delegates || [];
      const delMatch = isAll || isSlotAll || effDels.some(d => d.toLowerCase() === selectedDelegate.toLowerCase());
      if (!delMatch) continue;

      const dtStart = venueToUtcIso(MEETING_BASE_DATE, ev.day_idx, ev.start_local, MEETING_TZ_OFFSET_MIN);
      const dtEnd = venueToUtcIso(MEETING_BASE_DATE, ev.day_idx, ev.end_local, MEETING_TZ_OFFSET_MIN);
      if (!dtStart || !dtEnd) continue;

      const safeKey = (ev.slot_key || ('slot-' + ev.day_idx + '-' + ev.quarter + '-' + ev.stream_num)).replace(/[^a-zA-Z0-9_-]/g, '_');
      const uid = 'sa2-' + MEETING_NUMBER + '-' + safeName + '-' + safeKey + '-' + dtStart + '@3gpp.tools';

      let delTag = '';
      if (isSlotAll) {
        delTag = ev.is_cov ? '[All Team (cov)] ' : '[All Team] ';
      } else if (effDels.length > 0) {
        delTag = '[' + effDels.join(', ') + (ev.is_cov ? ' (cov)' : '') + '] ';
      }

      const summaryPrefix = ev.is_drafting ? '🛠️ [Drafting] ' + delTag : '[Str ' + ev.stream_num + '] ' + delTag;
      const summary = summaryPrefix + (ev.title || '').substring(0, 60);

      const fullLocation = ev.location_note ? (ev.room + ' (' + ev.location_note + ')') : (ev.room || '');
      const descParts = [
        'Session: SA2 #' + MEETING_NUMBER + ' (' + TDOC_ID + ')',
        'Room: ' + fullLocation,
        'Convener: ' + (ev.convener || 'N/A'),
        'Scheduled: ' + ev.title
      ];
      if (ev.note) descParts.push('Note: ' + ev.note);

      calLines.push('BEGIN:VEVENT');
      calLines.push('UID:' + uid);
      calLines.push('DTSTAMP:' + nowUtc);
      calLines.push('DTSTART:' + dtStart);
      calLines.push('DTEND:' + dtEnd);
      calLines.push('SUMMARY:' + summary.replace(/[\\r\\n]+/g, ' '));
      calLines.push('LOCATION:' + fullLocation.replace(/[\\r\\n]+/g, ' '));
      calLines.push('DESCRIPTION:' + descParts.join('\\\\n').replace(/[\\r\\n]+/g, ' '));
      calLines.push('END:VEVENT');
      count++;
    }

    calLines.push('END:VCALENDAR');
    const icsContent = calLines.join('\\r\\n');

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SA2_' + MEETING_NUMBER + '_' + safeName + '_' + TDOC_ID + '.ics';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

"""
