# --- File: src/modules/meetings/core/session_plan/exporter/css_views.py ---
"""
Embedded CSS (part 3/3): Mobile Day Cards, Conflicts/Unassigned/Now views,
slide-in slot inspector drawer, mobile breakpoints and print stylesheet.
"""

CSS_VIEWS = """\
  /* Mobile Day Cards View Styles */
  .day-tabs {
    display: flex; gap: 6px; overflow-x: auto; padding: 8px 14px; background: #FFFFFF;
    border-bottom: 1px solid var(--border); -webkit-overflow-scrolling: touch;
  }
  .day-tab {
    padding: 6px 14px; border-radius: 999px; font-weight: 600; font-size: 13px;
    background: #F1F5F9; color: var(--subtext); border: none; cursor: pointer; white-space: nowrap;
  }
  .day-tab.active { background: var(--primary); color: #FFFFFF; }
  .cards-content { padding: 12px 14px 60px 14px; }
  .quarter-card {
    background: var(--card); border: 1px solid var(--border); border-radius: 10px;
    margin-bottom: 12px; overflow: hidden; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
  }
  .quarter-header {
    background: #F8FAFC; padding: 8px 12px; border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
  }
  .quarter-title { font-size: 13px; font-weight: 700; color: #1E293B; }
  .quarter-time { font-size: 12px; color: var(--subtext); }
  .cards-slot-row {
    padding: 10px 12px; border-bottom: 1px solid #F1F5F9; transition: background 0.15s ease;
  }
  .cards-slot-row.is-inspected {
    outline: 2px solid #2563EB; outline-offset: -2px; background: #EFF6FF;
  }
  .cards-slot-row:last-child { border-bottom: none; }
  .cards-slot-row.drafting { background: var(--drafting-bg); border-left: 4px solid var(--drafting); }
  .cards-slot-row.unassigned { background: var(--unassigned-bg); border-left: 4px solid var(--unassigned); }
  .cards-slot-row.not-followed { background: #F8FAFC; border-left: 4px dashed #94A3B8; opacity: 0.75; }
  .slot-header-line { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; margin-bottom: 4px; }
  .stream-badge {
    font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; text-transform: uppercase;
  }
  .stream-badge.str-1 { background: #EEF2FF; color: var(--stream1); }
  .stream-badge.str-2 { background: #ECFDF5; color: var(--stream2); }
  .stream-badge.str-3 { background: #FFFBEB; color: var(--stream3); }
  .room-pill { font-size: 11px; color: var(--subtext); font-weight: 500; }
  .sub-slot-badge {
    font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 4px; background: #F1F5F9; color: #334155;
    border: 1px solid #CBD5E1; font-variant-numeric: tabular-nums;
  }
  .sub-slot-badge.is-half {
    background: #EFF6FF; color: #1E40AF; border-color: #BFDBFE;
  }
  .slot-content { font-size: 13px; font-weight: 500; color: #1E293B; margin-top: 3px; }
  .chips-container { margin-top: 4px; display: flex; gap: 4px; flex-wrap: wrap; }
  .note-chip {
    display: block; margin-top: 5px; padding: 4px 8px; border-radius: 6px;
    background: #FEF3C7; color: #92400E; font-size: 11px; font-weight: 500;
  }
  .hidden { display: none !important; }

  /* Conflicts View & Unassigned AIs View & Full "Now" View Styles */
  .now-view-shell { padding: 14px 16px 60px 16px; max-width: 1380px; margin: 0 auto; }
  .now-section-title {
    font-size: 13.5px; font-weight: 700; color: #0F172A; margin: 14px 0 8px 0;
    display: flex; align-items: center; gap: 6px;
  }
  .now-streams-row {
    display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px;
  }
  @media (max-width: 960px) {
    .now-streams-row { grid-template-columns: 1fr; }
  }
  .now-stream-card {
    background: #FFFFFF; border: 1px solid var(--grid-border); border-radius: 8px;
    padding: 11px 12px; display: flex; flex-direction: column; gap: 7px;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
  }
  .now-stream-card.str-1 { border-top: 4px solid var(--stream1); }
  .now-stream-card.str-2 { border-top: 4px solid var(--stream2); }
  .now-stream-card.str-3 { border-top: 4px solid var(--stream3); }
  .now-stream-card.is-idle { background: #F8FAFC; border-style: dashed; border-top-width: 1px; }
  .now-card-head { display: flex; justify-content: space-between; align-items: center; gap: 6px; flex-wrap: wrap; }
  .now-room-box {
    background: #F8FAFC; border: 1px solid var(--border); border-radius: 6px;
    padding: 6px 8px; display: flex; flex-direction: column; gap: 4px;
  }
  .now-room-title { font-size: 12.5px; font-weight: 700; color: #0F172A; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .now-loc-badge {
    display: inline-block; background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE;
    border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: 600;
  }
  .now-swap-badge {
    display: inline-block; background: #FFFBEB; color: #92400E; border: 1px solid #FDE68A;
    border-radius: 4px; padding: 1px 5px; font-size: 10px; font-weight: 700;
  }
  .now-room-meta { font-size: 10.5px; color: #64748B; }
  .now-topic-text { font-size: 12.5px; font-weight: 600; color: #0F172A; line-height: 1.38; }
  .now-tu2-hint {
    margin-top: auto; padding-top: 5px; border-top: 1px dashed var(--grid-border);
    font-size: 11px; color: #475569;
  }
  .now-upcoming-card {
    background: #FFFFFF; border: 1px solid var(--border); border-radius: 8px;
    padding: 10px 12px; margin-bottom: 8px; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
  }
  .now-upcoming-hdr {
    font-size: 12px; font-weight: 700; color: #1E293B; margin-bottom: 7px;
    display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px;
  }
  .now-upcoming-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px;
  }
  .now-upcoming-sub {
    background: #F8FAFC; border: 1px solid var(--border); border-radius: 6px; padding: 6px 8px;
  }
  .now-timeline-row {
    background: #F8FAFC; border: 1px solid var(--border); border-radius: 6px;
    padding: 7px 10px; margin-bottom: 6px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
  }
  .aux-view-container {
    padding: 14px 16px 60px 16px; max-width: 1360px; margin: 0 auto;
  }
  .aux-header-card {
    background: #FFFFFF; border: 1px solid var(--border); border-radius: 10px;
    padding: 12px 16px; margin-bottom: 12px; display: flex; justify-content: space-between;
    align-items: center; flex-wrap: wrap; gap: 10px; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
  }
  .aux-title { font-size: 15px; font-weight: 700; color: #0F172A; margin: 0; }
  .aux-subtitle { font-size: 12px; color: #64748B; margin-top: 2px; }
  .aux-table {
    width: 100%; border-collapse: separate; border-spacing: 0; background: #FFFFFF;
    border: 1px solid var(--grid-border); border-radius: 8px; overflow: hidden;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
  }
  .aux-table th {
    background: #F1F5F9; color: #1E293B; font-weight: 700; text-align: left;
    padding: 9px 12px; font-size: 11.5px; border-bottom: 1px solid var(--grid-border);
  }
  .aux-table td {
    padding: 9px 12px; font-size: 12px; border-bottom: 1px solid var(--border); vertical-align: top;
  }
  .aux-table tr:last-child td { border-bottom: none; }
  .aux-table tr:hover td { background: #F8FAFC; }
  .clash-stream-box {
    background: #FFFBEB; border: 1px solid #FDE68A; border-left: 3px solid #F59E0B; border-radius: 6px; padding: 6px 8px; margin-bottom: 4px;
  }
  .free-colleague-pill {
    display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 999px;
    font-size: 11px; font-weight: 700; margin: 2px 3px 2px 0;
  }
  .btn-jump-grid {
    padding: 4px 9px; font-size: 11px; font-weight: 600; border-radius: 5px;
    background: #EBF3FC; color: #1E5C99; border: 1px solid #BFD7F0; cursor: pointer;
    white-space: nowrap;
  }
  .btn-jump-grid:hover { background: #DFEDFA; }
  .empty-state-box {
    background: #FFFFFF; border: 1px dashed #CBD5E1; border-radius: 10px;
    padding: 32px 20px; text-align: center; color: #475569; font-size: 13px;
  }

  /* Click-Activated Slide-In Slot Inspector Drawer */
  .slot-inspector-drawer {
    position: fixed; top: 0; right: 0; bottom: 0; z-index: 9999;
    width: 390px; max-width: 92vw;
    background: #FFFFFF; border-left: 1px solid #CBD5E1;
    box-shadow: -8px 0 28px rgba(15, 23, 42, 0.14);
    display: flex; flex-direction: column;
    transform: translateX(100%);
    transition: transform 0.2s cubic-bezier(0.16, 1, 0.3, 1);
    pointer-events: none; visibility: hidden;
  }
  .slot-inspector-drawer.open {
    transform: translateX(0); pointer-events: auto; visibility: visible;
  }
  .drawer-header {
    display: flex; align-items: center; justify-content: space-between; gap: 10px;
    padding: 11px 14px; background: #F8FAFC; border-bottom: 1px solid #E2E8F0;
    flex-shrink: 0;
  }
  .drawer-heading-group { display: flex; flex-direction: column; gap: 1px; min-width: 0; }
  .drawer-kicker { font-size: 12px; font-weight: 700; color: #0F172A; letter-spacing: 0.01em; }
  .drawer-hint { font-size: 10.5px; color: var(--subtext); }
  .drawer-close-btn {
    border: 1px solid #CBD5E1; background: #FFFFFF; color: #475569;
    width: 28px; height: 28px; border-radius: 6px; font-size: 13px; font-weight: 700;
    cursor: pointer; display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .drawer-close-btn:hover { background: #F1F5F9; color: #0F172A; border-color: #94A3B8; }
  .drawer-body {
    padding: 14px 16px; overflow-y: auto; flex: 1 1 auto;
    font-family: inherit; font-size: 12.5px; line-height: 1.45; color: #0F172A;
    user-select: text;
  }
  .tt-header { font-weight: 700; font-size: 14px; color: #0F172A; margin-bottom: 4px; }
  .tt-time { color: var(--subtext); font-size: 12px; margin-bottom: 3px; }
  .tt-meta { color: var(--subtext); font-size: 11.5px; margin-bottom: 4px; }
  .tt-sep { border: none; border-top: 1px solid #E2E8F0; margin: 8px 0; }
  .tt-body { font-weight: 600; font-size: 12.5px; margin-bottom: 6px; }
  .tt-ais { font-size: 12px; color: #334155; line-height: 1.5; margin-top: 4px; }
  .tt-section { margin-top: 6px; font-size: 12px; }
  .tt-note { background: #FEF3C7; color: #92400E; padding: 6px 8px; border-radius: 5px; }

  @media (max-width: 768px) {
    header { position: static; padding: 6px 8px; }
    .day-tabs { position: sticky; top: 0; z-index: 90; }
    .overview-grid-container { margin: 2px 0 8px 0; max-height: 78vh; }
    .corner-hdr, .row-hdr { width: 108px; min-width: 108px; }
    .row-hdr { font-size: 9.5px; padding: 4px 5px !important; }
    .hdr-time-primary { font-size: 10.5px; }
    .overview-table { min-width: 1200px; }
    .overview-table.filtered-days { min-width: 340px; }
    .slot-inspector-drawer {
      top: auto; left: 0; right: 0; bottom: 0;
      width: 100%; max-width: 100%; max-height: 62vh;
      border-left: none; border-top: 1px solid #CBD5E1;
      border-radius: 14px 14px 0 0;
      box-shadow: 0 -8px 28px rgba(15, 23, 42, 0.18);
      transform: translateY(100%);
    }
    .slot-inspector-drawer.open { transform: translateY(0); }
  }

  /* Print Stylesheet */
  @media print {
    header, .del-strip, .filter-controls-row, .controls-bar, .view-switcher, .day-tabs, .now-dispatch-panel, .slot-inspector-drawer { display: none !important; }
    #nowView, #cardsView, #conflictsView, #unassignedView { display: none !important; }
    #overviewView { display: block !important; padding: 0 !important; }
    .overview-grid-container { max-height: none !important; overflow: visible !important; border: none !important; margin: 0 !important; }
    .overview-table { width: 100% !important; min-width: 0 !important; font-size: 8.5px !important; box-shadow: none !important; }
    .overview-table th, .overview-table td { padding: 3px 4px !important; position: static !important; }
    .slot-cell, .slot-card { opacity: 1 !important; filter: none !important; box-shadow: none !important; }
    @page { size: landscape; margin: 8mm; }
  }
"""
