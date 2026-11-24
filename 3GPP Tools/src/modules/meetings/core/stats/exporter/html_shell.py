"""HTML/CSS/JS shell renderer for the multi-scope meeting analytics dashboard."""

from __future__ import annotations

_DASHBOARD_CSS = """                    body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }
                    h1 { color: #333; text-align: center; margin-bottom: 10px; }
                    .selector-container { display: flex; justify-content: center; margin-bottom: 30px; background: #FFF; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #E0E0E0; }
                    .selector-container label { font-weight: bold; margin-right: 12px; align-self: center; color: #444; }
                    select { padding: 8px 16px; border-radius: 6px; border: 1px solid #CCCCCC; font-size: 14px; font-weight: bold; color: #005A9E; outline: none; background: #F4F8FC; cursor: pointer; }
                    select:hover { border-color: #005A9E; background: #EBF3FC; }
                    .kpi-container { display: flex; justify-content: center; gap: 14px; margin-bottom: 28px; flex-wrap: wrap; }
                    .kpi-card { background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.06); padding: 16px 14px; text-align: center; flex: 1 1 160px; min-width: 155px; max-width: 215px; border-top: 4px solid #1E5C99; border-left: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0; transition: transform 0.15s ease, box-shadow 0.15s ease; }
                    .kpi-card:hover { transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.09); }
                    .kpi-card h3 { margin: 0; font-size: 24px; color: #1E5C99; font-weight: 700; line-height: 1.2; }
                    .kpi-card p { margin: 4px 0 0; color: #475569; font-size: 11px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; }
                    .kpi-subtext { margin-top: 6px; font-size: 11px; color: #64748B; font-weight: 500; }
                    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
                    .chart-card { position: relative; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 40px 15px 15px 15px; height: 500px; display: flex; flex-direction: column; transition: all 0.3s ease; }
                    .chart-card > .chart-pane, .chart-card > .data-pane { flex-grow: 1; width: 100%; height: 100%; min-height: 0; }
                    .fs-btn { position: absolute; top: 10px; right: 10px; z-index: 100; cursor: pointer; background: #E1F0FF; color: #005A9E; border: 1px solid #99C9FF; border-radius: 4px; padding: 5px 10px; font-weight: bold; font-size: 12px; }
                    .fs-btn:hover { background: #CCE4FF; }
                    .view-toggle { position:absolute; top:10px; right:105px; z-index:100; display:inline-flex; width:auto !important; height:auto !important; flex-grow:0 !important; gap:0; border-radius:4px; overflow:hidden; }
                    .view-btn { cursor:pointer; background:#F8FAFC; color:#475569; border:1px solid #CBD5E1; padding:5px 9px; font-weight:bold; font-size:11px; line-height:16px; white-space:nowrap; }
                    .view-btn + .view-btn { border-left:0; }
                    .view-btn.active { background:#E1F0FF; color:#005A9E; border-color:#99C9FF; }
                    .chart-pane, .data-pane { flex-grow:1; width:100%; height:100%; min-height:0; }
                    .data-pane { display:none; overflow:auto; padding-top:8px; box-sizing:border-box; }
                    .data-table-wrap { overflow:auto; max-height:100%; }
                    .data-table { width:100%; border-collapse:collapse; font-size:12px; }
                    .data-table th { position:sticky; top:0; background:#F1F5F9; color:#334155; text-align:left; padding:8px; border-bottom:2px solid #CBD5E1; z-index:2; }
                    .data-table td { padding:7px 8px; border-bottom:1px solid #E2E8F0; vertical-align:top; }
                    .data-table tr:nth-child(even) { background:#F8FAFC; }
                    .chart-card.fullscreen { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 9999; margin: 0; border-radius: 0; padding: 50px 20px 20px 20px; box-sizing: border-box; }
                    .factions-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; }
                    .faction-box { background: #F9F9F9; border: 1px solid #E0E0E0; padding: 12px 15px; border-radius: 4px; }
                    .faction-box h4 { margin: 0 0 8px 0; color: #333; font-size: 15px; }
                    .faction-box p { margin: 0; font-size: 13px; color: #555; line-height: 1.5; }
                    .info-title-container { position: absolute; top: 15px; left: 15px; z-index: 50; pointer-events: none; }
                    .custom-tooltip { pointer-events: auto; position: relative; display: inline-block; cursor: help; color: #005A9E; font-size: 16px; margin-left: 8px; }
                    .custom-tooltip .custom-tooltip-text { visibility: hidden; width: 320px; background-color: #333; color: #fff; text-align: left; border-radius: 6px; padding: 15px; font-size: 13px; position: absolute; z-index: 1000; bottom: 125%; left: -10px; opacity: 0; transition: opacity 0.3s; box-shadow: 0 4px 8px rgba(0,0,0,0.2); line-height: 1.4; }
                    .custom-tooltip:hover .custom-tooltip-text { visibility: visible; opacity: 1; }
                    .export-data-link { display:inline-block; padding:7px 11px; border:1px solid #99C9FF; border-radius:5px; background:#E1F0FF; color:#005A9E; text-decoration:none; font-weight:bold; font-size:12px; }
                    .interactive-filter-bar { display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin:-15px auto 22px; max-width:1200px; padding:10px 12px; background:#FFFFFF; border:1px solid #E2E8F0; border-radius:7px; box-shadow:0 2px 4px rgba(0,0,0,.04); }
                    .interactive-filter-label { font-size:12px; font-weight:bold; color:#475569; }
                    .filter-chip { display:inline-flex; align-items:center; gap:5px; padding:4px 8px; border-radius:14px; background:#E1F0FF; border:1px solid #99C9FF; color:#005A9E; font-size:11px; font-weight:bold; }
                    .filter-chip button { border:0; background:transparent; cursor:pointer; color:#005A9E; font-weight:bold; padding:0; }
                    .clear-crossfilters { border:1px solid #CBD5E1; background:#F8FAFC; color:#475569; border-radius:4px; padding:4px 8px; cursor:pointer; font-size:11px; }
                    .crossfilter-summary { margin-left:auto; color:#64748B; font-size:11px; }
                    .data-tools { display:flex; gap:7px; align-items:center; margin:0 0 8px 0; flex-wrap:wrap; }
                    .data-search { flex:1 1 220px; min-width:160px; border:1px solid #CBD5E1; border-radius:4px; padding:6px 8px; font-size:12px; }
                    .data-tools button { cursor:pointer; border:1px solid #CBD5E1; background:#F8FAFC; color:#334155; border-radius:4px; padding:5px 8px; font-size:11px; font-weight:bold; }
                    .data-row-count { color:#64748B; font-size:11px; white-space:nowrap; }
                    .data-table th { cursor:pointer; user-select:none; }
                    .sort-indicator { color:#94A3B8; font-size:10px; }
                    .data-table tr.table-filter-hidden { display:none; }
                    .plotly-click-hint { text-align:center; color:#64748B; font-size:11px; margin-top:-4px; }
                    .filter-chip.inactive { opacity:.45; border-style:dashed; }
                    .matching-explorer { max-width:1200px; margin:-12px auto 24px; background:#FFF; border:1px solid #E2E8F0; border-radius:7px; padding:10px 12px; }
                    .matching-explorer summary { cursor:pointer; font-weight:bold; color:#334155; }
                    .matching-tools { display:flex; gap:7px; align-items:center; margin:10px 0 8px; flex-wrap:wrap; }
                    .matching-tools input { flex:1 1 260px; min-width:180px; border:1px solid #CBD5E1; border-radius:4px; padding:6px 8px; }
                    .matching-tools button { cursor:pointer; border:1px solid #CBD5E1; background:#F8FAFC; border-radius:4px; padding:5px 8px; font-size:11px; font-weight:bold; }
"""

_DASHBOARD_JS = r"""                    function switchAI(selectedId) {
                        const sections = document.querySelectorAll('.dashboard-view-panel');
                        sections.forEach(sec => { sec.style.display = 'none'; });

                        const activeSec = document.getElementById(selectedId);
                        if(activeSec) {
                            activeSec.style.display = 'block';
                            window.dispatchEvent(new Event('resize'));
                            renderDashboardFilters();
                        }
                    }
                    function toggleChartData(btn, mode) {
                        const card = btn.closest('.chart-card');
                        if(!card) return;
                        const chart = card.querySelector('.chart-pane');
                        const data = card.querySelector('.data-pane');
                        if(!chart || !data) return;
                        chart.style.display = mode === 'chart' ? 'block' : 'none';
                        data.style.display = mode === 'data' ? 'block' : 'none';
                        card.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
                        btn.classList.add('active');
                        if(mode === 'chart') setTimeout(() => window.dispatchEvent(new Event('resize')), 30);
                    }

                    function toggleFullscreen(btn) {
                        const card = btn.parentElement;
                        card.classList.toggle('fullscreen');
                        btn.innerHTML = card.classList.contains('fullscreen') ? '✖ Close' : '⛶ Expand';
                        setTimeout(() => { window.dispatchEvent(new Event('resize')); }, 50);
                    }

                    const dashboardFilters = {company: new Set(), ai: new Set(), status: new Set()};
                    let dashboardRows = [];
                    let matchingSearch = '';
                    let matchingShowAll = false;
                    const MATCHING_PREVIEW_LIMIT = 100;

                    function currentScopeInfo() {
                        const select = document.getElementById('scope-select');
                        const selected = select?.value || 'global';
                        const option = select?.selectedOptions?.[0];
                        return {
                            id: selected,
                            wi: option?.dataset?.wi || '',
                            ai: option?.dataset?.ai || '',
                            isAiScope: selected.startsWith('ai_')
                        };
                    }

                    function currentScopeRows() {
                        const scope = currentScopeInfo();
                        let rows = dashboardRows;
                        if(scope.wi) rows = rows.filter(r => (r.wis || []).includes(scope.wi));
                        if(scope.ai) rows = rows.filter(r => r.ai === scope.ai);
                        return rows;
                    }

                    function filteredRows() {
                        const scope = currentScopeInfo();
                        return currentScopeRows().filter(r =>
                            (!dashboardFilters.company.size || (r.companies || []).some(c => dashboardFilters.company.has(c))) &&
                            (scope.isAiScope || !dashboardFilters.ai.size || dashboardFilters.ai.has(r.ai)) &&
                            (!dashboardFilters.status.size || dashboardFilters.status.has(r.status))
                        );
                    }

                    function toggleDashboardFilter(dimension, value) {
                        value = String(value ?? '').trim();
                        if(!value || !dashboardFilters[dimension]) return;
                        const scope = currentScopeInfo();
                        if(dimension === 'ai' && scope.isAiScope) return;
                        const set = dashboardFilters[dimension];
                        set.has(value) ? set.delete(value) : set.add(value);
                        matchingShowAll = false;
                        renderDashboardFilters();
                    }

                    function removeDashboardFilter(dimension, value) {
                        dashboardFilters[dimension]?.delete(String(value));
                        matchingShowAll = false;
                        renderDashboardFilters();
                    }

                    function clearDashboardFilters() {
                        Object.values(dashboardFilters).forEach(s => s.clear());
                        matchingSearch = '';
                        matchingShowAll = false;
                        const input = document.getElementById('matching-search');
                        if(input) input.value = '';
                        renderDashboardFilters();
                    }

                    function renderDashboardFilters() {
                        const host = document.getElementById('active-crossfilters');
                        if(!host) return;
                        host.innerHTML = '';
                        const scope = currentScopeInfo();
                        const names = {company:'Company', ai:'AI', status:'Status'};
                        Object.entries(dashboardFilters).forEach(([dim, values]) => {
                            values.forEach(value => {
                                const chip = document.createElement('span');
                                chip.className = 'filter-chip' + (dim === 'ai' && scope.isAiScope ? ' inactive' : '');
                                if(dim === 'ai' && scope.isAiScope) chip.title = 'AI filters are temporarily ignored because the selected scope already fixes the Agenda Item.';
                                chip.append(document.createTextNode(`${names[dim]}: ${value}`));
                                const b=document.createElement('button'); b.textContent='×';
                                b.onclick=()=>removeDashboardFilter(dim,value); chip.appendChild(b);
                                host.appendChild(chip);
                            });
                        });
                        const scopeRows = currentScopeRows();
                        const rows = filteredRows();
                        const summary=document.getElementById('crossfilter-summary');
                        if(summary) summary.textContent=`${rows.length} / ${scopeRows.length} TDocs match`;
                        renderMatchingTDocs();
                    }

                    function handleDashboardPlotClick(ev) {
                        const pt=ev?.points?.[0]; if(!pt) return;
                        const dim=pt.data?.meta?.filterDimension; if(!dim || !dashboardFilters[dim]) return;
                        let value='';
                        if(Array.isArray(pt.customdata)) value=pt.customdata[0];
                        else if(pt.customdata != null) value=pt.customdata;
                        else if(dim==='company') value=pt.y;
                        else if(dim==='status') value=pt.label;
                        toggleDashboardFilter(dim, value);
                    }

                    function initializeDashboard() {
                        const raw=document.getElementById('dashboard-filter-data');
                        try { dashboardRows=JSON.parse(raw?.textContent || '[]'); } catch(e) { dashboardRows=[]; }
                        document.querySelectorAll('.plotly-graph-div').forEach(plot => {
                            if(plot.dataset.exploreWired || typeof plot.on !== 'function' || !Array.isArray(plot.data)) return;
                            const interactive = plot.data.some(trace => dashboardFilters[trace?.meta?.filterDimension]);
                            if(!interactive) return;
                            plot.on('plotly_click', handleDashboardPlotClick);
                            plot.dataset.exploreWired='1';
                            const pane=plot.closest('.chart-pane');
                            const card=plot.closest('.chart-card');
                            if(pane && card && !card.querySelector('.plotly-click-hint')) {
                                const hint=document.createElement('div');
                                hint.className='plotly-click-hint';
                                hint.textContent='Click a bar, slice or point to explore matching TDocs.';
                                pane.appendChild(hint);
                            }
                        });
                        document.querySelectorAll('.data-search').forEach(filterStatsTable);
                        renderDashboardFilters();
                    }

                    function matchingRowsAfterSearch() {
                        const q=(matchingSearch||'').trim().toLowerCase();
                        const rows=filteredRows();
                        if(!q) return rows;
                        return rows.filter(r => [r.tdoc,r.type,r.ai,r.aiDisplay,(r.wis||[]).join(' '),r.source,(r.companies||[]).join(' '),r.status]
                            .some(v => String(v||'').toLowerCase().includes(q)));
                    }

                    function escapeHtml(value) {
                        return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
                    }

                    function renderMatchingTDocs() {
                        const body=document.getElementById('matching-tdocs-body');
                        const count=document.getElementById('matching-row-count');
                        if(!body) return;
                        const rows=matchingRowsAfterSearch();
                        const shown=matchingShowAll ? rows : rows.slice(0, MATCHING_PREVIEW_LIMIT);
                        body.innerHTML=shown.map((r,i)=>`<tr data-original-index="${i}"><td data-value="${escapeHtml(r.tdoc)}">${escapeHtml(r.tdoc)}</td><td data-value="${escapeHtml(r.type || '')}">${escapeHtml(r.type || '')}</td><td data-value="${escapeHtml(r.ai)}">${escapeHtml(r.ai)}</td><td data-value="${escapeHtml(r.aiDisplay)}">${escapeHtml(r.aiDisplay)}</td><td data-value="${escapeHtml((r.wis||[]).join('; '))}">${escapeHtml((r.wis||[]).join('; '))}</td><td data-value="${escapeHtml(r.source || (r.companies||[]).join(', '))}">${escapeHtml(r.source || (r.companies||[]).join(', '))}</td><td data-value="${escapeHtml(r.status)}">${escapeHtml(r.status)}</td></tr>`).join('');
                        if(count) count.textContent = shown.length < rows.length ? `Showing ${shown.length} of ${rows.length}` : `${rows.length} rows`;
                        const more=document.getElementById('matching-show-all');
                        if(more) {
                            more.style.display = rows.length > MATCHING_PREVIEW_LIMIT ? 'inline-block' : 'none';
                            more.textContent = matchingShowAll ? 'Show first 100' : 'Show all';
                        }
                    }

                    function filterMatchingTDocs(input) {
                        matchingSearch=input.value || '';
                        matchingShowAll=false;
                        renderMatchingTDocs();
                    }
                    function toggleMatchingShowAll() { matchingShowAll=!matchingShowAll; renderMatchingTDocs(); }

                    function refreshTableVisibility(table, query) {
                        query=(query||'').toLowerCase();
                        let visible=0, total=0;
                        Array.from(table.tBodies[0]?.rows || []).forEach(row => {
                            total++;
                            const show=!query || row.textContent.toLowerCase().includes(query);
                            row.classList.toggle('table-filter-hidden', !show);
                            if(show) visible++;
                        });
                        const count=table.closest('.data-pane')?.querySelector('.data-row-count');
                        if(count) count.textContent=`${visible} / ${total} rows`;
                    }
                    function filterStatsTable(input) {
                        const table=input.closest('.data-pane')?.querySelector('.data-table');
                        if(table) refreshTableVisibility(table,input.value);
                    }
                    function parseSortValue(v) {
                        const cleaned=String(v).trim().replace(/[% ,]/g,'');
                        return cleaned!=='' && !Number.isNaN(Number(cleaned)) ? Number(cleaned) : String(v).toLowerCase();
                    }
                    function sortStatsTable(th) {
                        const table=th.closest('table'), body=table?.tBodies[0]; if(!body) return;
                        const idx=Array.from(th.parentNode.children).indexOf(th);
                        const asc=th.dataset.sort !== 'asc';
                        table.querySelectorAll('th').forEach(h=>{h.dataset.sort=''; const i=h.querySelector('.sort-indicator'); if(i)i.textContent='↕';});
                        th.dataset.sort=asc?'asc':'desc'; const indicator=th.querySelector('.sort-indicator'); if(indicator)indicator.textContent=asc?'↑':'↓';
                        const rows=Array.from(body.rows);
                        rows.sort((a,b)=>{
                            const av=parseSortValue(a.cells[idx]?.dataset.value ?? a.cells[idx]?.textContent ?? '');
                            const bv=parseSortValue(b.cells[idx]?.dataset.value ?? b.cells[idx]?.textContent ?? '');
                            const cmp=(typeof av==='number' && typeof bv==='number') ? av-bv : String(av).localeCompare(String(bv),undefined,{numeric:true,sensitivity:'base'});
                            return asc?cmp:-cmp;
                        });
                        rows.forEach(r=>body.appendChild(r));
                    }
                    function resetStatsTable(btn) {
                        const pane=btn.closest('.data-pane'); if(!pane) return;
                        const input=pane.querySelector('.data-search'); if(input) input.value='';
                        const table=pane.querySelector('.data-table'), body=table?.tBodies[0];
                        if(body) Array.from(body.rows).sort((a,b)=>Number(a.dataset.originalIndex||0)-Number(b.dataset.originalIndex||0)).forEach(r=>body.appendChild(r));
                        table?.querySelectorAll('th').forEach(h=>{h.dataset.sort=''; const i=h.querySelector('.sort-indicator'); if(i)i.textContent='↕';});
                        if(table) refreshTableVisibility(table,'');
                    }
                    function visibleTableMatrix(btn) {
                        const pane=btn.closest('.data-pane'), table=pane?.querySelector('.data-table'); if(!table)return [];
                        const headers=Array.from(table.tHead?.rows[0]?.cells||[]).map(c=>c.textContent.replace(/[↕↑↓]/g,'').trim());
                        const rows=Array.from(table.tBodies[0]?.rows||[]).filter(r=>!r.classList.contains('table-filter-hidden'))
                            .map(r=>Array.from(r.cells).map(c=>c.dataset.value ?? c.textContent.trim()));
                        return [headers,...rows];
                    }
                    function csvEscape(v){const s=String(v??''); return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;}
                    function copyVisibleTable(btn) {
                        const m=visibleTableMatrix(btn); if(!m.length)return;
                        navigator.clipboard?.writeText(m.map(r=>r.join('\t')).join('\n'));
                        const old=btn.textContent; btn.textContent='✓ Copied'; setTimeout(()=>btn.textContent=old,900);
                    }
                    function downloadVisibleTableCsv(btn) {
                        const m=visibleTableMatrix(btn); if(!m.length)return;
                        const csv='\uFEFF'+m.map(r=>r.map(csvEscape).join(',')).join('\r\n');
                        const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'}), url=URL.createObjectURL(blob);
                        const a=document.createElement('a'); a.href=url;
                        const scope=document.getElementById('scope-select')?.selectedOptions?.[0]?.textContent?.trim() || 'Statistics';
                        a.download=(scope.replace(/[^a-z0-9_-]+/gi,'_') || 'Statistics')+'_table.csv'; a.click(); URL.revokeObjectURL(url);
                    }
                    function matchingMatrix() {
                        const rows=matchingRowsAfterSearch();
                        return [['TDoc','Type','AI','Topic','WI/SI','Source','Status'], ...rows.map(r=>[r.tdoc,r.type || '',r.ai,r.aiDisplay,(r.wis||[]).join('; '),r.source || (r.companies||[]).join(', '),r.status])];
                    }
                    function copyMatchingTDocs(btn) {
                        const m=matchingMatrix(); navigator.clipboard?.writeText(m.map(r=>r.join('\t')).join('\n'));
                        const old=btn.textContent; btn.textContent='✓ Copied'; setTimeout(()=>btn.textContent=old,900);
                    }
                    function downloadMatchingTDocsCsv() {
                        const m=matchingMatrix(); const csv='\uFEFF'+m.map(r=>r.map(csvEscape).join(',')).join('\r\n');
                        const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'}), url=URL.createObjectURL(blob);
                        const a=document.createElement('a'); a.href=url; a.download='Matching_TDocs.csv'; a.click(); URL.revokeObjectURL(url);
                    }"""


def render_dashboard_html(
    meeting_name: str,
    dropdown_options: list[str],
    interactive_json: str,
    views_html_buffer: list[str],
    plotly_js_version: str,
) -> str:
    """Assemble the complete standalone HTML analytics dashboard document."""
    return f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>3GPP Multi-Scope Statistics - {meeting_name}</title>
                <script src="https://cdn.plot.ly/plotly-{plotly_js_version}.min.js"></script>
                <style>
{_DASHBOARD_CSS}
                </style>
                <script>
{_DASHBOARD_JS}
                </script>
            </head>
            <body>
                <div style="text-align:right;margin-bottom:8px;"><a class="export-data-link" href="Statistics_Data.xlsx">📥 Statistics Data (Excel)</a></div><h1>📊 {meeting_name} - TDoc Analytics Dashboard</h1>

                <div class="selector-container">
                    <label for="scope-select">🎯 Analytics Scope / Agenda Item:</label>
                    <select id="scope-select" onchange="switchAI(this.value)">
                        {" ".join(dropdown_options)}
                    </select>
                </div>
                <div class="interactive-filter-bar">
                    <span class="interactive-filter-label">🔎 Explore data:</span>
                    <span id="active-crossfilters"></span>
                    <button type="button" class="clear-crossfilters" onclick="clearDashboardFilters()">Clear filters</button>
                    <span id="crossfilter-summary" class="crossfilter-summary"></span>
                </div>
                <script type="application/json" id="dashboard-filter-data">{interactive_json}</script>
                <details class="matching-explorer" id="matching-explorer">
                    <summary>Matching TDocs</summary>
                    <div class="matching-tools">
                        <input id="matching-search" type="search" placeholder="Search matching TDocs..." oninput="filterMatchingTDocs(this)">
                        <span id="matching-row-count" class="data-row-count"></span>
                        <button id="matching-show-all" type="button" onclick="toggleMatchingShowAll()">Show all</button>
                        <button type="button" onclick="copyMatchingTDocs(this)">📋 Copy</button>
                        <button type="button" onclick="downloadMatchingTDocsCsv()">⬇ CSV</button>
                    </div>
                    <div class="data-table-wrap" style="max-height:420px;">
                        <table class="data-table" id="matching-tdocs-table"><thead><tr>
                            <th onclick="sortStatsTable(this)">TDoc <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">Type <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">AI <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">Topic <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">WI/SI <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">Source <span class="sort-indicator">↕</span></th>
                            <th onclick="sortStatsTable(this)">Status <span class="sort-indicator">↕</span></th>
                        </tr></thead><tbody id="matching-tdocs-body"></tbody></table>
                    </div>
                </details>

                {" ".join(views_html_buffer)}
                <script>
                    window.addEventListener('load', initializeDashboard, {{once:true}});
                </script>
            </body>
            </html>
            """
