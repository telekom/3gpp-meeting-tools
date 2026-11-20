# --- File: src/modules/meetings/core/session_plan/exporter/css_layout.py ---
"""
Embedded CSS (part 1/3): root tokens, body/viewport layout, sticky header,
NOW dispatch panel, delegate/day filter strips and search bar.
"""

CSS_LAYOUT = """
  :root {
    --bg: #F1F5F9;
    --surface: #F8FAFC;
    --card: #FFFFFF;
    --text: #0F172A;
    --text-strong: #1E293B;
    --subtext: #475569;
    --muted: #64748B;
    --primary: #1E5C99;
    --primary-bright: #2563EB;
    --primary-light: #EBF3FC;
    --stream1: #4F46E5;
    --stream2: #059669;
    --stream3: #D97706;
    --drafting: #DC2626;
    --drafting-bg: #FEF2F2;
    --drafting-bdr: #FECACA;
    --unassigned: #D97706;
    --unassigned-bg: #FFFBEB;
    --border: #E2E8F0;
    --grid-border: #CBD5E1;
    --day-hdr-h: 27px;
  }
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  body {
    margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
    font-size: 12px; line-height: 1.35;
  }
  #overviewView {
    display: flex; flex-direction: column; min-height: 0; padding: 4px 6px 8px 6px;
  }
  @media (min-width: 769px) {
    body {
      height: 100vh; display: flex; flex-direction: column; overflow: hidden;
    }
    #overviewView {
      flex: 1 1 auto; min-height: 0; display: flex; flex-direction: column; padding: 6px 10px 8px 10px;
    }
    #nowView, #cardsView, #conflictsView, #unassignedView {
      flex: 1 1 auto; min-height: 0; overflow-y: auto;
    }
  }
  header {
    position: sticky; top: 0; z-index: 100; background: #FFFFFF; flex-shrink: 0;
    border-bottom: 1px solid var(--border); padding: 5px 12px;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
  }
  .header-top {
    display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; flex-wrap: wrap; gap: 6px;
  }
  .title-block { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .title { font-size: 14.5px; font-weight: 700; color: var(--text-strong); margin: 0; letter-spacing: -0.01em; }
  .badge-tdoc { font-size: 11px; background: #EEF2FF; color: #3730A3; border: 1px solid #C7D2FE; padding: 1px 7px; border-radius: 999px; font-weight: 700; }
  .badge-snap { font-size: 10.5px; background: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; padding: 1px 6px; border-radius: 999px; font-weight: 600; }
  .badge-utc {
    font-size: 10.5px; font-weight: 700; background: #ECFDF5; color: #065F46;
    border: 1px solid #A7F3D0; padding: 1px 7px; border-radius: 999px; white-space: nowrap;
  }
  .badge-live-clock {
    display: inline-flex; align-items: center; gap: 4px;
    font-size: 11px; font-weight: 600; font-variant-numeric: tabular-nums;
    background: #EFF6FF; color: #1E3A8A; border: 1px solid #BFDBFE;
    padding: 2px 9px; border-radius: 999px; white-space: nowrap;
  }
  .badge-live-clock strong { font-size: 11.5px; font-weight: 700; }
  .header-actions { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .btn-cal-download {
    padding: 3px 9px; font-size: 11px; font-weight: 600; border-radius: 6px;
    background: #1E5C99; color: #FFFFFF; border: 1px solid #15426E; cursor: pointer;
    transition: all 0.15s ease; display: inline-flex; align-items: center; gap: 4px; white-space: nowrap;
  }
  .btn-cal-download:hover { background: #15426E; }
  .badge-clash {
    font-size: 11px; font-weight: 700; background: #FFFBEB; color: #B45309;
    border: 1px solid #F59E0B; padding: 2px 8px; border-radius: 999px;
    display: inline-flex; align-items: center; gap: 4px; cursor: pointer;
    transition: all 0.15s ease; white-space: nowrap;
  }
  .badge-clash:hover { background: #FEF3C7; border-color: #D97706; }
  .live-banner {
    display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11px; font-weight: 600;
    padding: 4px 10px; border-radius: 6px; background: #ECFDF5; color: #065F46;
    border: 1px solid #A7F3D0; margin-top: 2px; cursor: pointer; flex-shrink: 0;
  }
  .live-banner:hover { background: #D1FAE5; }
  .live-banner-right { display: inline-flex; align-items: center; gap: 8px; flex-shrink: 0; }
  .live-banner-close {
    border: none; background: transparent; color: #065F46; font-weight: 700; font-size: 12px;
    cursor: pointer; padding: 0 4px; border-radius: 4px; line-height: 1.2;
  }
  .live-banner-close:hover { background: rgba(6, 95, 70, 0.12); }
  .pulse-dot {
    width: 7px; height: 7px; border-radius: 50%; background: #10B981;
    display: inline-block; animation: pulseLive 1.8s infinite;
  }
  @keyframes pulseLive {
    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
    70% { transform: scale(1.1); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
  }
  .grid-cell.has-clash .slot-card {
    background: #FFFBEB !important;
    border: 1px solid #F59E0B !important;
    border-left: 4px solid #F59E0B !important;
  }
  .grid-cell.has-clash .cell-title {
    color: #92400E; font-weight: 500;
  }
  .grid-cell.has-clash .is-clash-pill {
    display: inline-flex !important;
  }
  .del-pill-badge.clash-pill-ring {
    border: 1.5px solid #EF4444 !important;
    box-shadow: 0 0 0 1px rgba(239, 68, 68, 0.22);
  }
  .cards-slot-row.has-clash {
    border-left: 4px solid #F59E0B !important; background: #FFFBEB !important;
  }
  .grid-cell.is-live-now .slot-card, .grid-cell.is-now-focused .slot-card {
    background: #ECFDF5 !important; border: 2px solid #10B981 !important;
    box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.20) !important;
  }
  .cards-slot-row.is-live-now, .cards-slot-row.is-now-focused {
    border-left: 4px solid #10B981 !important; background: #ECFDF5 !important;
  }
  .view-switcher {
    display: inline-flex; background: var(--surface); padding: 2px; border-radius: 7px; border: 1px solid var(--border); flex-wrap: wrap; gap: 2px;
  }
  .view-btn {
    padding: 4px 10px; font-size: 11.5px; font-weight: 600; border: 1px solid transparent; background: transparent;
    color: var(--subtext); border-radius: 5px; cursor: pointer; transition: all 0.15s ease; white-space: nowrap;
  }
  .view-btn:hover:not(.active) { background: #E2E8F0; color: var(--text); }
  .view-btn.active {
    background: #1E5C99; color: #FFFFFF; border-color: #15426E; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.12);
  }
  .view-btn.has-alert { color: #B91C1C; }
  .view-btn.has-alert.active { background: #DC2626; color: #FFFFFF; border-color: #B91C1C; }

  /* Now Dispatch Panel ("Highlight Now" - Where are we now & where should each delegate go?) */
  .now-dispatch-panel {
    background: #F0FDF4; border: 1px solid #86EFAC; border-radius: 8px;
    padding: 10px 12px; margin: 5px 0; box-shadow: 0 2px 6px rgba(0,0,0,0.05);
  }
  .now-panel-header {
    display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;
    border-bottom: 1px solid #BBF7D0; padding-bottom: 8px; margin-bottom: 8px;
  }
  .now-panel-title { font-size: 13px; font-weight: 700; color: #065F46; display: flex; align-items: center; gap: 6px; }
  .now-slot-controls { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .now-select {
    padding: 3px 8px; font-size: 11.5px; font-weight: 600; border-radius: 5px;
    border: 1px solid #86EFAC; background: #FFFFFF; color: #064E3B; cursor: pointer;
  }
  .now-step-btn {
    padding: 3px 8px; font-size: 11px; font-weight: 700; border-radius: 5px;
    border: 1px solid #86EFAC; background: #FFFFFF; color: #065F46; cursor: pointer;
  }
  .now-step-btn:hover { background: #DCFCE7; }
  .now-dispatch-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 8px;
  }
  .now-del-card {
    background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 7px 10px;
    display: flex; flex-direction: column; gap: 3px; cursor: pointer; transition: transform 0.1s ease, box-shadow 0.1s ease;
  }
  .now-del-card:hover { transform: translateY(-1px); box-shadow: 0 2px 6px rgba(0,0,0,0.07); }
  .now-del-card.status-busy { border-left: 4px solid #2563EB; }
  .now-del-card.status-clash { border-left: 4px solid #F59E0B; background: #FFFBEB; }
  .now-del-card.status-free { border-left: 4px solid #10B981; background: #F8FAFC; }
  .now-del-top { display: flex; justify-content: space-between; align-items: center; gap: 4px; }
  .now-del-name { font-weight: 700; font-size: 12px; }
  .now-del-status-tag { font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 999px; }
  .now-del-dest { font-size: 11px; color: #1E293B; font-weight: 600; }
  .now-del-topic { font-size: 10.5px; color: #475569; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .now-streams-summary {
    margin-top: 8px; padding-top: 6px; border-top: 1px dashed #BBF7D0;
    display: flex; gap: 8px; flex-wrap: wrap; font-size: 11px; color: #065F46;
  }

  /* Delegate Selector & Day Filter Strip */
  .controls-bar {
    display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px 10px;
    padding-top: 2px;
  }
  .del-strip {
    display: flex; gap: 4px; align-items: center; overflow-x: auto; padding: 1px 0;
    -webkit-overflow-scrolling: touch; flex-wrap: wrap;
  }
  .strip-label { font-size: 10px; font-weight: 700; color: #64748B; letter-spacing: 0.02em; margin-right: 1px; }
  .del-filter-btn {
    padding: 2px 8px; font-size: 10.5px; font-weight: 600; border-radius: 999px; cursor: pointer;
    background: #F1F5F9; color: #334155; border: 1px solid var(--border); white-space: nowrap;
    transition: all 0.15s ease;
  }
  .del-filter-btn[data-del="ALL"].active {
    background: #1E5C99; color: #FFFFFF; border-color: #15426E;
  }
  .del-filter-btn[data-del]:not([data-del="ALL"]) {
    background: var(--del-bg, #F1F5F9); color: var(--del-color, #334155); border-color: var(--del-bdr, var(--border));
  }
  .del-filter-btn[data-del]:not([data-del="ALL"]).active {
    background: var(--del-color, #2563EB); color: #FFFFFF; border-color: var(--del-color, #2563EB);
    box-shadow: 0 1px 4px rgba(0,0,0,0.14);
  }
  .strip-sep { width: 1px; height: 15px; background: var(--border); margin: 0 2px; }
  .btn-toggle { background: var(--surface); }
  #btnUnassigned.active { background: #D97706; color: #FFFFFF; border-color: #B45309; }
  #btnDrafting.active { background: #DC2626; color: #FFFFFF; border-color: #DC2626; }

  .day-filter-strip {
    display: inline-flex; gap: 3px; align-items: center; flex-wrap: wrap;
  }
  .day-filter-btn {
    padding: 2px 7px; font-size: 10.5px; font-weight: 600; border-radius: 5px; cursor: pointer;
    background: var(--surface); color: var(--subtext); border: 1px solid var(--grid-border);
    white-space: nowrap; transition: all 0.12s ease;
  }
  .day-filter-btn:hover:not(.active) { background: #E2E8F0; color: var(--text); }
  .day-filter-btn.active {
    background: #1E5C99; color: #FFFFFF; border-color: #15426E;
  }
  .day-filter-btn[data-day-filter="TODAY_PLUS"].active {
    background: #059669; border-color: #047857;
  }

  /* Filter Controls Row */
  .filter-controls-row {
    display: flex; gap: 6px; align-items: center; flex-wrap: wrap; flex: 1 1 360px; justify-content: flex-end;
  }
  .search-box {
    flex: 1 1 160px; max-width: 240px; min-width: 140px; padding: 3px 8px; border: 1px solid var(--grid-border); border-radius: 6px;
    font-size: 11.5px; outline: none; background: var(--surface); color: var(--text);
  }
  .search-box:focus { background: #FFFFFF; border-color: var(--primary); box-shadow: 0 0 0 2px rgba(30, 92, 153, 0.12); }
  .grid-row-controls { display: flex; gap: 4px; align-items: center; font-size: 10.5px; flex-wrap: wrap; }
  .toggle-label {
    display: inline-flex; align-items: center; gap: 3px; cursor: pointer; color: var(--subtext); font-weight: 600;
    background: var(--surface); border: 1px solid var(--border); padding: 2px 6px; border-radius: 5px; white-space: nowrap;
  }
  .toggle-label:hover { background: #F1F5F9; }
  .toggle-label input { cursor: pointer; margin: 0; }
  .str-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }
  .str-dot.s1 { background: var(--stream1); }
  .str-dot.s2 { background: var(--stream2); }
  .str-dot.s3 { background: var(--stream3); }

"""
