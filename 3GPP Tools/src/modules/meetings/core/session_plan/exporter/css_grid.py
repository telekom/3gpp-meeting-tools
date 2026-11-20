# --- File: src/modules/meetings/core/session_plan/exporter/css_grid.py ---
"""
Embedded CSS (part 2/3): Overview Grid table, sticky day/period headers,
slot cards, delegate dimming and cell contents.
"""

CSS_GRID = """\
  /* Overview Grid View Styles (Full-Viewport Container with Locked Day Header Row & Period Column) */
  .overview-grid-container {
    margin: 6px 10px 12px 10px; overflow: auto; flex: 1 1 auto; min-height: 300px;
    -webkit-overflow-scrolling: touch; border: 1px solid #94A3B8; border-radius: 8px;
    background: #FFFFFF; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
  }
  @media (min-width: 769px) {
    .overview-grid-container { margin: 0; }
  }
  .overview-table {
    border-collapse: separate; border-spacing: 0; width: 100%; min-width: 920px;
    background: #FFFFFF; table-layout: fixed;
  }
  .overview-table.filtered-days { min-width: 0; }
  .overview-table th, .overview-table td {
    border-right: 1px solid var(--border); border-bottom: 1px solid var(--border);
    padding: 4px 6px; vertical-align: top; font-size: 11px;
  }
  .overview-table td.grid-cell {
    padding: 3px; background: #FFFFFF;
  }
  .overview-table td.grid-cell.day-odd { background: #F8FAFC; }
  .overview-table td.grid-cell.day-today { background: #F0FDF4; }
  .overview-table th.tu-split-col,
  .overview-table td.tu-split-col {
    border-right: 1px dashed #CBD5E1;
  }
  .overview-table th.day-end-col,
  .overview-table td.day-end-col {
    border-right: 2px solid #94A3B8 !important;
  }
  .overview-table tr.period-end-row th,
  .overview-table tr.period-end-row td {
    border-bottom: 2px solid #94A3B8;
  }
  .day-hidden { display: none !important; }
  .overview-table th:last-child, .overview-table td:last-child { border-right: none; }
  .overview-table tr:last-child td { border-bottom: none; }

  .hdr-days-row th {
    background: #F8FAFC; color: #0F172A; font-weight: 700; text-align: center;
    position: sticky; top: 0; z-index: 20; border-top: 2px solid #64748B; border-bottom: 1px solid var(--grid-border);
    box-shadow: 0 1px 0 var(--grid-border); padding: 4px 6px; white-space: nowrap;
  }
  .hdr-days-row th.col-day-hdr.day-odd {
    background: #EEF2F6; border-top-color: #94A3B8;
  }
  .hdr-days-row th.col-day-hdr.day-today {
    background: #ECFDF5; color: #065F46; border-top: 3px solid #10B981;
  }
  .corner-hdr {
    background: #E2E8F0 !important; text-align: left !important; font-size: 11px;
    position: sticky; top: 0; left: 0; z-index: 30 !important; width: 152px; min-width: 152px;
    border-top: none !important; border-right: 2px solid #94A3B8 !important;
  }
  .hdr-sub { font-size: 9px; font-weight: 500; color: #64748B; }
  .hdr-date { font-size: 10.5px; font-weight: 600; color: #475569; }
  .badge-today {
    display: inline-block; font-size: 9px; font-weight: 700; background: #DCFCE7;
    color: #15803D; border: 1px solid #86EFAC; padding: 0px 5px; border-radius: 4px; margin-left: 4px;
  }
  .hdr-tu-row th {
    background: #FFFFFF; color: #64748B; font-size: 9.5px; font-weight: 600; text-align: center;
    position: sticky; top: var(--day-hdr-h, 25px); z-index: 19; padding: 2px 4px;
    border-bottom: 2px solid #CBD5E1; box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    white-space: nowrap; min-width: 0;
  }
  .hdr-tu-row th.day-odd { background: #F1F5F9; }
  .hdr-tu-row th.day-today { background: #F0FDF4; color: #166534; }
  .row-hdr {
    background: #F8FAFC; color: #1E293B; font-weight: 600; text-align: left;
    position: sticky; left: 0; z-index: 10; width: 152px; min-width: 152px; font-size: 10.5px;
    padding: 4px 6px !important; border-right: 2px solid #94A3B8 !important;
  }
  .hdr-time-line {
    display: flex; align-items: center; justify-content: space-between; gap: 4px;
  }
  .hdr-time-primary {
    font-size: 11.5px; font-weight: 700; color: #0F172A; font-variant-numeric: tabular-nums;
  }
  .hdr-now-tag {
    display: none; font-size: 8.5px; font-weight: 700; padding: 0px 4px; border-radius: 4px;
    background: #DCFCE7; color: #166534; border: 1px solid #86EFAC;
  }
  tr.grid-row.is-live-row-hdr th.row-hdr { background: #F0FDF4; }
  tr.grid-row.is-live-row-hdr .hdr-time-primary,
  tr.grid-row.is-live-row-hdr .hdr-break-title { color: #166534; }
  tr.grid-row.is-live-row-hdr .hdr-now-tag { display: inline-block; }
  .hdr-stream-line { display: flex; align-items: center; gap: 5px; margin-top: 2px; }
  .hdr-stream-sub { font-size: 10px; font-weight: 500; color: #475569; }
  .hdr-time-local { font-size: 9.5px; font-weight: 600; color: #64748B; }
  .hdr-break-title { font-size: 11px; font-weight: 700; color: #92400E; }
  .hdr-time-break {
    font-size: 10.5px; font-weight: 700; color: #B45309; margin-top: 2px; font-variant-numeric: tabular-nums;
  }
  .hdr-room { font-size: 9.5px; font-weight: 500; color: #64748B; }
  .hdr-times { font-size: 9.5px; font-weight: 500; color: #64748B; }
  .stream-hdr-s1 { border-left: 4px solid var(--stream1); }
  .stream-hdr-s2 { border-left: 4px solid var(--stream2); }
  .stream-hdr-s3 { border-left: 4px solid var(--stream3); }
  .break-hdr { background: #FFFBEB; font-weight: 600; color: #92400E; border-left: 4px solid #D97706; }

  /* Grid Cells & Unified Rounded Slot Cards */
  .grid-cell {
    background: #FFFFFF; transition: opacity 0.2s ease, filter 0.2s ease;
    cursor: default; position: relative;
  }
  .slot-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-left: 3px solid #CBD5E1;
    border-radius: 5px;
    padding: 4px 6px;
    min-height: 44px;
    height: 100%;
    display: flex;
    justify-content: space-between;
    align-items: stretch;
    gap: 5px;
    transition: box-shadow 0.15s ease, border-color 0.15s ease;
  }
  .slot-card-body {
    flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; gap: 2px;
  }
  .grid-cell.slot-cell, .cards-slot-row { cursor: pointer; }
  .grid-cell.slot-cell:hover .slot-card {
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.08);
    border-color: #94A3B8;
  }
  .grid-cell.slot-cell.is-inspected .slot-card {
    outline: 2px solid #2563EB;
    outline-offset: -1px;
    box-shadow: 0 3px 10px rgba(37, 99, 235, 0.18);
  }
  .grid-cell.empty-slot-cell {
    text-align: center; color: #CBD5E1; font-weight: 500; vertical-align: middle !important;
  }
  .grid-cell.empty-break-cell {
    text-align: center; color: #CBD5E1; font-weight: 500; background: #FAFAFA; vertical-align: middle !important;
  }
  .grid-cell.empty-slot-cell.is-empty-drafting {
    color: #94A3B8; font-weight: 500;
  }
  .grid-cell.is-empty-drafting .slot-card {
    background: #F8FAFC !important; border: 1px dashed #E2E8F0 !important;
  }
  .grid-cell.is-empty-drafting .cell-title {
    color: #94A3B8; font-weight: 400;
  }
  .grid-cell.is-drafting .slot-card {
    background: var(--drafting-bg) !important;
    border: 1px solid var(--drafting-bdr) !important;
    border-left: 4px solid var(--drafting) !important;
  }
  .grid-cell.is-drafting .cell-title {
    color: #991B1B; font-weight: 500;
  }
  body.filter-unassigned .grid-cell.is-unassigned .slot-card {
    background: #FFFBEB !important; border: 1px solid #F59E0B !important;
  }
  .grid-cell.is-not-followed .slot-card {
    background: #F8FAFC !important;
    border: 1px dashed #CBD5E1 !important;
    opacity: 0.72;
  }
  .grid-cell.is-not-followed .cell-title {
    color: #64748B;
  }

  /* Delegate Dimming / Highlighting */
  .grid-cell.dimmed {
    opacity: 0.22; filter: grayscale(85%);
  }
  .grid-cell.highlighted .slot-card {
    box-shadow: 0 0 0 2px var(--primary);
  }

  /* Cell Contents */
  .cell-del-float {
    display: flex; flex-direction: column; align-items: flex-end; gap: 2px;
    flex-shrink: 0;
  }
  .cell-meta-row {
    display: flex; align-items: center; gap: 4px; flex-wrap: wrap; margin-bottom: 2px;
  }
  .cell-time-pill {
    display: inline-block; font-size: 9px; font-weight: 700; font-variant-numeric: tabular-nums;
    border-radius: 3px; padding: 0px 4px; white-space: nowrap;
  }
  .cell-time-pill.is-full {
    background: #F8FAFC; color: #475569; border: 1px solid #E2E8F0; font-weight: 600;
  }
  .cell-time-pill.is-half {
    background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE;
  }
  .cell-time-pill.is-custom-time {
    background: #FFFBEB; color: #92400E; border: 1px solid #FDE68A;
  }
  .cell-room-row { margin-bottom: 1px; }
  .cell-room-override {
    display: inline-block; font-size: 9px; font-weight: 700; color: #475569;
    background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 3px; padding: 0px 4px;
    max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  .cell-top-row { display: flex; gap: 4px; align-items: center; margin-bottom: 3px; flex-wrap: wrap; }
  .cell-badge {
    font-size: 9px; font-weight: 700; padding: 1px 4px; border-radius: 3px; text-transform: uppercase;
  }
  .badge-str.str-1 { background: #EEF2FF; color: var(--stream1); }
  .badge-str.str-2 { background: #ECFDF5; color: var(--stream2); }
  .badge-str.str-3 { background: #FFFBEB; color: var(--stream3); }
  .badge-drafting { background: #FEE2E2; color: var(--drafting); }
  .badge-tu { background: #E2E8F0; color: #334155; }
  .cell-room { font-size: 10px; color: #64748B; font-weight: 500; }
  .cell-title { font-size: 10.5px; font-weight: 500; color: #0F172A; line-height: 1.3; overflow-wrap: break-word; word-break: normal; }
  .cell-chips-row { display: flex; gap: 3px; flex-wrap: wrap; margin-top: 2px; }
  .ai-chip {
    font-size: 10px; font-weight: 600; padding: 1px 5px; border-radius: 3px; background: #E0E7FF; color: #3730A3;
  }
  .ai-chip.unassigned-chip {
    background: #FEF3C7; color: #B45309; border: 1px dashed #F59E0B;
  }
  .ai-chip.not-followed-chip {
    background: #F1F5F9; color: #64748B; border: 1px solid #CBD5E1;
  }
  .chip-count { font-size: 9px; font-weight: 700; color: #475569; margin-left: 2px; }
  .cell-del-row { display: flex; gap: 3px; flex-wrap: wrap; margin-top: 2px; }
  .del-pill-badge {
    font-size: 8.5px; font-weight: 700; padding: 1px 5px; border-radius: 7px;
    display: inline-flex; align-items: center; line-height: 1.3; white-space: nowrap;
  }
  .del-pill-badge.is-all {
    background: #EEF2FF; color: #3730A3; border: 1px solid #6366F1;
  }
  .del-pill-badge.is-more-pill {
    background: #F1F5F9; color: #334155; border: 1px solid #CBD5E1;
  }
  .del-pill-badge.is-unassigned-pill {
    background: #FEF3C7; color: #B45309; border: 1px dashed #F59E0B;
  }
  .del-pill-badge.is-not-followed-pill {
    background: #F1F5F9; color: #64748B; border: 1px solid #CBD5E1;
  }
  .del-pill-badge.is-clash-pill {
    display: none; background: #FEF3C7; color: #B45309; border: 1px solid #F59E0B; font-size: 9px; padding: 0px 4px;
  }
  .cell-note-pill {
    margin-top: 2px; font-size: 9px; font-weight: 500; color: #92400E; background: #FEF3C7;
    padding: 1px 4px; border-radius: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }

"""
