"""Standalone HTML and PDF print stylesheet for compiled 3GPP Technical Dossiers."""

from __future__ import annotations

DOSSIER_CSS = """
:root {
  --primary: #005A9E;
  --primary-hover: #004578;
  --primary-light: #EFF6FF;
  --bg-page: #F8FAFC;
  --bg-card: #FFFFFF;
  --text-main: #1E293B;
  --text-muted: #64748B;
  --border-color: #E2E8F0;
  --border-subtle: #CBD5E1;
  --diff-ins-bg: #DCFCE7;
  --diff-ins-text: #166534;
  --diff-del-bg: #FEE2E2;
  --diff-del-text: #991B1B;
}

* {
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 14px;
  line-height: 1.6;
  color: var(--text-main);
  background-color: var(--bg-page);
  margin: 0;
  padding: 0;
}

/* --- Container & Layout --- */
.dossier-wrapper {
  max-width: 1100px;
  margin: 0 auto;
  padding: 24px 20px 80px 20px;
}

/* --- Floating Action Navigation Bar --- */
.floating-nav {
  position: sticky;
  top: 0;
  z-index: 1000;
  background: rgba(255, 255, 255, 0.95);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--border-color);
  padding: 10px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.04);
}

.floating-nav .title-pill {
  font-weight: 700;
  color: var(--primary);
  font-size: 13px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.floating-nav .nav-controls {
  display: flex;
  align-items: center;
  gap: 8px;
}

.floating-nav select {
  font-size: 12px;
  padding: 5px 10px;
  border: 1px solid var(--border-subtle);
  border-radius: 4px;
  background-color: #fff;
  color: var(--text-main);
  outline: none;
  max-width: 320px;
}

.floating-nav button, .floating-nav a.btn {
  font-size: 12px;
  font-weight: 600;
  padding: 5px 12px;
  border-radius: 4px;
  border: 1px solid var(--border-subtle);
  background-color: #fff;
  color: var(--text-main);
  text-decoration: none;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  transition: all 0.15s ease-in-out;
}

.floating-nav button:hover, .floating-nav a.btn:hover {
  background-color: var(--primary-light);
  border-color: var(--primary);
  color: var(--primary);
}

.floating-nav button.primary {
  background-color: var(--primary);
  color: #fff;
  border-color: var(--primary);
}

.floating-nav button.primary:hover {
  background-color: var(--primary-hover);
}

/* --- Header & Summary Cover Card --- */
.cover-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 28px 32px;
  margin-bottom: 28px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.cover-card h1 {
  margin: 0 0 8px 0;
  color: var(--primary);
  font-size: 24px;
  font-weight: 700;
}

.cover-card .meta-subtitle {
  color: var(--text-muted);
  font-size: 13px;
  margin-bottom: 20px;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 12px;
  margin: 16px 0;
}

.stat-chip {
  background: var(--primary-light);
  border: 1px solid #BFDBFE;
  border-radius: 6px;
  padding: 10px 14px;
  text-align: center;
}

.stat-chip .val {
  font-size: 20px;
  font-weight: 700;
  color: var(--primary);
}

.stat-chip .lbl {
  font-size: 11px;
  text-transform: uppercase;
  color: var(--text-muted);
  font-weight: 600;
}

/* --- Master Index Table --- */
.index-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 24px;
  margin-bottom: 32px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.index-card h2 {
  margin-top: 0;
  color: var(--text-main);
  font-size: 18px;
  border-bottom: 2px solid var(--primary-light);
  padding-bottom: 8px;
}

table.index-table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 14px;
  font-size: 12px;
}

table.index-table th {
  background: #F1F5F9;
  color: #334155;
  text-align: left;
  padding: 8px 10px;
  border: 1px solid var(--border-color);
  font-weight: 600;
}

table.index-table td {
  padding: 8px 10px;
  border: 1px solid var(--border-color);
  vertical-align: top;
}

table.index-table tr:nth-child(even) {
  background-color: #F8FAFC;
}

table.index-table tr:hover {
  background-color: var(--primary-light);
}

.badge {
  display: inline-block;
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 11px;
  font-weight: 600;
  text-align: center;
}

.badge-cr { background: #FEF3C7; color: #92400E; }
.badge-pcr { background: #E0E7FF; color: #3730A3; }
.badge-ls { background: #E0F2FE; color: #075985; }
.badge-disc { background: #F1F5F9; color: #475569; }
.badge-agreed { background: #DCFCE7; color: #166534; }
.badge-noted { background: #F3F4F6; color: #374151; }

/* --- TDoc Contribution Sections --- */
.tdoc-section {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 28px 32px;
  margin-bottom: 36px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.tdoc-header {
  border-bottom: 2px solid var(--border-color);
  padding-bottom: 14px;
  margin-bottom: 20px;
}

.tdoc-header-top {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 8px;
}

.tdoc-id-title {
  font-size: 20px;
  font-weight: 700;
  color: var(--primary);
  margin: 0;
}

.tdoc-actions {
  display: flex;
  gap: 8px;
}

.tdoc-actions a {
  font-size: 11px;
  font-weight: 600;
  padding: 4px 8px;
  border-radius: 4px;
  border: 1px solid var(--border-subtle);
  background: #F8FAFC;
  color: var(--text-main);
  text-decoration: none;
}

.tdoc-actions a:hover {
  background: var(--primary-light);
  color: var(--primary);
  border-color: var(--primary);
}

.tdoc-meta-row {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 12px;
  color: var(--text-muted);
  margin-top: 6px;
}

.tdoc-meta-row strong {
  color: var(--text-main);
}

/* --- Document Body & CommonMark Rendering --- */
.tdoc-body {
  font-size: 13.5px;
  line-height: 1.6;
}

.tdoc-body h1, .tdoc-body h2, .tdoc-body h3, .tdoc-body h4 {
  color: #0F172A;
  margin-top: 24px;
  margin-bottom: 10px;
  font-weight: 600;
}

.tdoc-body h1 { font-size: 18px; border-bottom: 1px solid var(--border-color); padding-bottom: 4px; }
.tdoc-body h2 { font-size: 16px; }
.tdoc-body h3 { font-size: 14.5px; }

.tdoc-body table {
  width: 100%;
  border-collapse: collapse;
  margin: 16px 0;
  font-size: 12px;
}

.tdoc-body th, .tdoc-body td {
  border: 1px solid var(--border-color);
  padding: 6px 10px;
  vertical-align: top;
}

.tdoc-body th {
  background-color: #F8FAFC;
  font-weight: 600;
  color: #334155;
}

.tdoc-body img {
  max-width: 100%;
  height: auto;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  margin: 14px 0;
  display: block;
}

.tdoc-body pre, .tdoc-body code {
  font-family: Consolas, Monaco, "Courier New", monospace;
  font-size: 12px;
  background-color: #F1F5F9;
  border-radius: 4px;
}

.tdoc-body code {
  padding: 2px 4px;
}

.tdoc-body pre {
  padding: 12px;
  overflow-x: auto;
  border: 1px solid var(--border-color);
}

/* --- Revision Marks & Diff Highlighting --- */
.diff-ins, ins {
  background-color: var(--diff-ins-bg);
  color: var(--diff-ins-text);
  text-decoration: underline;
  padding: 1px 3px;
  border-radius: 2px;
}

.diff-del, del {
  background-color: var(--diff-del-bg);
  color: var(--diff-del-text);
  text-decoration: line-through;
  padding: 1px 3px;
  border-radius: 2px;
}

.change-marker {
  display: block;
  background-color: var(--primary-light);
  color: var(--primary);
  border-left: 4px solid var(--primary);
  font-weight: 700;
  font-size: 12px;
  padding: 6px 12px;
  margin: 16px 0;
  border-radius: 0 4px 4px 0;
}

/* --- Print Styling for Vector PDF Export --- */
@media print {
  body {
    background-color: #FFFFFF !important;
    font-size: 10pt !important;
    color: #000000 !important;
  }
  .dossier-wrapper {
    max-width: 100% !important;
    padding: 0 !important;
  }
  .no-print, .floating-nav, .tdoc-actions {
    display: none !important;
  }
  .cover-card, .index-card, .tdoc-section {
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
  }
  .tdoc-section {
    page-break-before: always !important;
  }
  .tdoc-section:first-of-type {
    page-break-before: avoid !important;
  }
  table.index-table th {
    background-color: #EEEEEE !important;
    color: #000000 !important;
  }
  .change-marker {
    background-color: #F0F4F8 !important;
    border-left: 3px solid #005A9E !important;
    color: #005A9E !important;
  }
  table {
    page-break-inside: auto;
  }
  tr {
    page-break-inside: avoid;
    page-break-after: auto;
  }
  @page {
    margin: 15mm 15mm 20mm 15mm;
  }
}
"""
