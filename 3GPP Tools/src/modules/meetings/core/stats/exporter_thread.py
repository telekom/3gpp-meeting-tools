# --- File: src/modules/meetings/core/stats/exporter_thread.py ---
import os
import re
import html
import json
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
import plotly.express as px

from core.config.plot_styles import PALETTE, THEME_COLOR, CLUSTER_PALETTE
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.agenda_manager import AgendaManager
from .plot_agenda import generate_ai_volume_plot, generate_ai_status_plot
from .plot_status import generate_outcomes_plot
from .plot_contributors import generate_top_contributors_plot, generate_company_ai_heatmap
from .plot_alliances import compute_global_communities, generate_alliance_plots
from .plot_insights import (generate_revision_intensity_plot, generate_company_specialization_plot,
                            build_revision_stats, build_company_specialization)


class StatisticsExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, meeting_dir: Path, tdocs_data: list, mtg_info: dict, config: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.tdocs_data = tdocs_data
        self.mtg_info = mtg_info
        self.config = config
        self.export_dir = self.meeting_dir / "Export"

        self.cfg_resolution = self.config.get("resolution", 1.5)
        self.cfg_threshold = self.config.get("threshold", 1)
        self.cfg_top_count = self.config.get("top_count", 30)
        self.cfg_export_html = self.config.get("export_html_plots", False)

        # ---> RETRIEVE NEW SETTINGS
        self.cfg_hm_companies = self.config.get("heatmap_top_companies", 25)
        self.cfg_hm_ais = self.config.get("heatmap_top_ais", 25)

        self.THEME_COLOR = THEME_COLOR
        self.PALETTE = PALETTE
        self.CLUSTER_PALETTE = CLUSTER_PALETTE

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            plots_dir = self.export_dir / "Interactive_Plots"
            if self.cfg_export_html:
                plots_dir.mkdir(parents=True, exist_ok=True)

            df = pd.DataFrame(self.tdocs_data)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            df = df[
                ~df['TDoc Status'].astype(str).str.lower().str.contains(
                    'withdrawn',
                    na=False
                )
            ].copy().reset_index(drop=True)

            df['Clean_Companies'] = df['Source'].apply(
                CompanySanitizer.get_matching_contributors
            )

            # Optional semantic AI metadata from the meeting's local agenda.csv.
            # The raw Agenda Item remains the canonical grouping/filter key.
            agenda_map = AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")
            ai_acronym_counts = {}
            for ai in df['Agenda Item'].fillna('').astype(str).str.strip().unique():
                item = agenda_map.get(ai)
                if item and item.acronym:
                    ai_acronym_counts[item.acronym] = ai_acronym_counts.get(item.acronym, 0) + 1

            def ai_meta(ai_value):
                ai = str(ai_value or '').strip()
                item = agenda_map.get(ai)
                if not item:
                    return pd.Series({'AI_Acronym': '', 'AI_Topic': '', 'AI_Full_Description': '', 'AI_Display': ai, 'AI_Axis': ai})
                acronym, topic = item.acronym or '', item.topic_suffix or ''
                if acronym and ai_acronym_counts.get(acronym, 0) > 1 and topic:
                    display = f"{ai} ({acronym} – {topic})"
                elif acronym:
                    display = f"{ai} ({acronym})"
                else:
                    display = ai
                # Keep plot axes compact. Rich semantic metadata remains in hover/data tables.
                if acronym and ai_acronym_counts.get(acronym, 0) > 1 and topic:
                    short_topic = topic if len(topic) <= 24 else topic[:21].rstrip() + '...'
                    axis = f"{ai}<br>{short_topic}"
                elif acronym:
                    axis = f"{ai}<br>{acronym}"
                else:
                    axis = ai
                return pd.Series({'AI_Acronym': acronym, 'AI_Topic': topic,
                                  'AI_Full_Description': item.description, 'AI_Display': display, 'AI_Axis': axis})

            ai_meta_df = df['Agenda Item'].apply(ai_meta)
            for col in ai_meta_df.columns:
                df[col] = ai_meta_df[col]

            def extract_wi_list(value):
                text = str(value or '').strip()
                if not text or text.lower() in {'none', 'nan', '-', 'unknown', 'dummy'}:
                    return []
                return [part.strip() for part in re.split(r'[,;\n]+', text) if part.strip()]

            if 'Related WIs' in df.columns:
                df['Clean_WI_List'] = df['Related WIs'].apply(extract_wi_list)
            else:
                df['Clean_WI_List'] = [[] for _ in range(len(df))]

            # Compact browser-side dataset used for interactive cross-filtering.
            # Raw TDoc text is intentionally not embedded; only dimensions needed by the dashboard.
            interactive_rows = []
            for _, row in df.iterrows():
                interactive_rows.append({
                    'tdoc': str(row.get('TDoc', '') or ''),
                    'ai': str(row.get('Agenda Item', '') or '').strip(),
                    'aiDisplay': str(row.get('AI_Display', row.get('Agenda Item', '')) or '').strip(),
                    'wis': list(row.get('Clean_WI_List', []) or []),
                    'companies': list(row.get('Clean_Companies', []) or []),
                    'status': str(row.get('TDoc Status', '') or '').strip(),
                    'isRevision': bool(str(row.get('Is revision of', '') or '').strip()) or bool(
                        re.search(r'(?:r|rev)\\d{1,2}[a-zA-Z]?$', str(row.get('TDoc', '') or ''), re.IGNORECASE)
                    ),
                })
            interactive_json = json.dumps(interactive_rows, ensure_ascii=False).replace('</', '<\\/')

            global_factions = compute_global_communities(df, self.cfg_resolution)

            # --- GENERATE GLOBAL PLOTS ---
            g_html_ai, g_table_ai = generate_ai_volume_plot(df, plots_dir, self.THEME_COLOR, prefix_id="Global",
                                                save_html=self.cfg_export_html)
            g_html_ai_status, g_table_ai_status = generate_ai_status_plot(df, plots_dir, self.PALETTE, prefix_id="Global",
                                                       save_html=self.cfg_export_html)
            g_html_status, g_table_status = generate_outcomes_plot(df, plots_dir, self.PALETTE, prefix_id="Global",
                                                   save_html=self.cfg_export_html)
            g_html_comp, total_companies, g_table_comp = generate_top_contributors_plot(df, plots_dir, self.THEME_COLOR,
                                                                          self.cfg_top_count, prefix_id="Global",
                                                                          save_html=self.cfg_export_html)

            # ---> PASS DYNAMIC LIMITS TO HEATMAP
            g_html_heatmap, g_table_heatmap = generate_company_ai_heatmap(df, plots_dir, prefix_id="Global",
                                                         save_html=self.cfg_export_html,
                                                         top_comps_count=self.cfg_hm_companies,
                                                         top_ais_count=self.cfg_hm_ais)

            g_html_net, g_html_cluster, g_html_cohesion, g_html_list, g_alliance_tables = generate_alliance_plots(df, plots_dir,
                                                                                               self.cfg_threshold,
                                                                                               self.CLUSTER_PALETTE,
                                                                                               global_factions,
                                                                                               prefix_id="Global",
                                                                                               save_html=self.cfg_export_html)
            g_html_revision, g_table_revision = generate_revision_intensity_plot(
                df, plots_dir, self.THEME_COLOR, prefix_id="Global", save_html=self.cfg_export_html)
            g_html_specialization, g_table_specialization = generate_company_specialization_plot(
                df, plots_dir, self.THEME_COLOR, prefix_id="Global", save_html=self.cfg_export_html)

            raw_ais = df['Agenda Item'].dropna().unique()

            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

            unique_ais = sorted([str(ai).strip() for ai in raw_ais if str(ai).strip()], key=natural_sort_key)

            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"

            views_html_buffer = []
            dropdown_options = ['<option value="global">🌐 Overall Meeting View</option>']

            # Inject the Global plots into the template
            views_html_buffer.append(
                self._compile_view_block("global", len(df), total_companies,
                                         g_html_ai, g_html_status, g_html_comp,
                                         g_html_net, g_html_cluster, g_html_cohesion, g_html_list,
                                         ai_status_html=g_html_ai_status, heatmap_html=g_html_heatmap,
                                         ai_volume_table=g_table_ai, status_table=g_table_status,
                                         comp_table=g_table_comp, ai_status_table=g_table_ai_status,
                                         heatmap_table=g_table_heatmap, network_table=g_alliance_tables.get("network", ""),
                                         cluster_table=g_alliance_tables.get("cluster", ""), cohesion_table=g_alliance_tables.get("cohesion", ""),
                                         revision_html=g_html_revision, revision_table=g_table_revision,
                                         specialization_html=g_html_specialization, specialization_table=g_table_specialization,
                                         is_visible=True))

            unique_wis = sorted({wi for wis in df['Clean_WI_List'] for wi in wis}, key=lambda x: x.lower())
            if unique_wis:
                dropdown_options.append('<option disabled>──────── Work / Study Items ────────</option>')
            for wi_idx, wi_name in enumerate(unique_wis):
                wi_df = df[df['Clean_WI_List'].apply(lambda values, w=wi_name: w in values)].copy().reset_index(drop=True)
                if wi_df.empty:
                    continue
                safe_id = f"wi_{wi_idx}"
                safe_wi_prefix = "WI_" + re.sub(r'[\\/*?:"<>|]', '_', wi_name)
                dropdown_options.append(
                    f'<option value="{safe_id}" data-wi="{html.escape(wi_name, quote=True)}">🧩 {html.escape(wi_name)} ({len(wi_df)} TDocs)</option>')
                wi_html_ai, wi_table_ai = generate_ai_volume_plot(wi_df, plots_dir, self.THEME_COLOR, safe_wi_prefix, self.cfg_export_html)
                wi_html_status, wi_table_status = generate_outcomes_plot(wi_df, plots_dir, self.PALETTE, safe_wi_prefix, self.cfg_export_html)
                wi_html_comp, wi_companies, wi_table_comp = generate_top_contributors_plot(
                    wi_df, plots_dir, self.THEME_COLOR, self.cfg_top_count, safe_wi_prefix, self.cfg_export_html)
                wi_html_net, wi_html_cluster, wi_html_cohesion, wi_html_list, wi_alliance_tables = generate_alliance_plots(
                    wi_df, plots_dir, self.cfg_threshold, self.CLUSTER_PALETTE, global_factions, safe_wi_prefix, self.cfg_export_html)
                wi_html_revision, wi_table_revision = generate_revision_intensity_plot(
                    wi_df, plots_dir, self.THEME_COLOR, safe_wi_prefix, self.cfg_export_html)
                views_html_buffer.append(self._compile_view_block(
                    safe_id, len(wi_df), wi_companies, wi_html_ai, wi_html_status, wi_html_comp,
                    wi_html_net, wi_html_cluster, wi_html_cohesion, wi_html_list,
                    ai_volume_table=wi_table_ai, status_table=wi_table_status, comp_table=wi_table_comp,
                    network_table=wi_alliance_tables.get('network', ''), cluster_table=wi_alliance_tables.get('cluster', ''),
                    cohesion_table=wi_alliance_tables.get('cohesion', ''), revision_html=wi_html_revision,
                    revision_table=wi_table_revision, is_visible=False))

            if unique_ais:
                dropdown_options.append('<option disabled>──────── Agenda Items ────────</option>')

            for idx, ai_name in enumerate(unique_ais):
                ai_df = df[
                    df['Agenda Item'].astype(str).str.strip() == ai_name
                    ].copy().reset_index(drop=True)
                if ai_df.empty: continue

                safe_id = f"ai_{idx}"
                clean_ai_name = re.sub(r'[\\/*?:\"<>|]', '_', str(ai_name))
                safe_ai_prefix = "AI_" + clean_ai_name

                ai_display = str(ai_df['AI_Display'].iloc[0]) if 'AI_Display' in ai_df.columns else ai_name
                dropdown_options.append(
                    f'<option value="{safe_id}" data-ai="{html.escape(ai_name, quote=True)}">📌 {html.escape(ai_display)} ({len(ai_df)} TDocs)</option>')

                ai_html_status, ai_table_status = generate_outcomes_plot(ai_df, plots_dir, self.PALETTE, safe_ai_prefix,
                                                        save_html=self.cfg_export_html)
                ai_html_comp, ai_companies, ai_table_comp = generate_top_contributors_plot(ai_df, plots_dir, self.THEME_COLOR,
                                                                            self.cfg_top_count, safe_ai_prefix,
                                                                            save_html=self.cfg_export_html)
                ai_html_net, ai_html_cluster, ai_html_cohesion, ai_html_list, ai_alliance_tables = generate_alliance_plots(ai_df, plots_dir,
                                                                                                       self.cfg_threshold,
                                                                                                       self.CLUSTER_PALETTE,
                                                                                                       global_factions,
                                                                                                       safe_ai_prefix,
                                                                                                       save_html=self.cfg_export_html)

                # Keep AI Status and Heatmap explicitly None for individual Agenda Items
                views_html_buffer.append(self._compile_view_block(
                    safe_id, len(ai_df), ai_companies,
                    ai_volume_html=None, status_html=ai_html_status, comp_html=ai_html_comp,
                    net_html=ai_html_net, cluster_html=ai_html_cluster, cohesion_html=ai_html_cohesion,
                    list_html=ai_html_list,
                    ai_status_html=None, heatmap_html=None,
                    status_table=ai_table_status, comp_table=ai_table_comp,
                    network_table=ai_alliance_tables.get("network", ""), cluster_table=ai_alliance_tables.get("cluster", ""),
                    cohesion_table=ai_alliance_tables.get("cohesion", ""), is_visible=False
                ))

            dashboard_template = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>3GPP Multi-Scope Statistics - {meeting_name}</title>
                <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
                <style>
                    body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }}
                    h1 {{ color: #333; text-align: center; margin-bottom: 10px; }}
                    .selector-container {{ display: flex; justify-content: center; margin-bottom: 30px; background: #FFF; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #E0E0E0; }}
                    .selector-container label {{ font-weight: bold; margin-right: 12px; align-self: center; color: #444; }}
                    select {{ padding: 8px 16px; border-radius: 6px; border: 1px solid #CCCCCC; font-size: 14px; font-weight: bold; color: #005A9E; outline: none; background: #F4F8FC; cursor: pointer; }}
                    select:hover {{ border-color: #005A9E; background: #EBF3FC; }}
                    .kpi-container {{ display: flex; justify-content: center; gap: 20px; margin-bottom: 40px; }}
                    .kpi-card {{ background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 220px; border-top: 4px solid #005A9E; }}
                    .kpi-card h3 {{ margin: 0; font-size: 32px; color: #005A9E; }}
                    .kpi-card p {{ margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }}
                    .grid-container {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
                    .chart-card {{ position: relative; background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 40px 15px 15px 15px; height: 500px; display: flex; flex-direction: column; transition: all 0.3s ease; }}
                    .chart-card > .chart-pane, .chart-card > .data-pane {{ flex-grow: 1; width: 100%; height: 100%; min-height: 0; }}
                    .fs-btn {{ position: absolute; top: 10px; right: 10px; z-index: 100; cursor: pointer; background: #E1F0FF; color: #005A9E; border: 1px solid #99C9FF; border-radius: 4px; padding: 5px 10px; font-weight: bold; font-size: 12px; }}
                    .fs-btn:hover {{ background: #CCE4FF; }}
                    .view-toggle {{ position:absolute; top:10px; right:105px; z-index:100; display:inline-flex; width:auto !important; height:auto !important; flex-grow:0 !important; gap:0; border-radius:4px; overflow:hidden; }}
                    .view-btn {{ cursor:pointer; background:#F8FAFC; color:#475569; border:1px solid #CBD5E1; padding:5px 9px; font-weight:bold; font-size:11px; line-height:16px; white-space:nowrap; }}
                    .view-btn + .view-btn {{ border-left:0; }}
                    .view-btn.active {{ background:#E1F0FF; color:#005A9E; border-color:#99C9FF; }}
                    .chart-pane, .data-pane {{ flex-grow:1; width:100%; height:100%; min-height:0; }}
                    .data-pane {{ display:none; overflow:auto; padding-top:8px; box-sizing:border-box; }}
                    .data-table-wrap {{ overflow:auto; max-height:100%; }}
                    .data-table {{ width:100%; border-collapse:collapse; font-size:12px; }}
                    .data-table th {{ position:sticky; top:0; background:#F1F5F9; color:#334155; text-align:left; padding:8px; border-bottom:2px solid #CBD5E1; z-index:2; }}
                    .data-table td {{ padding:7px 8px; border-bottom:1px solid #E2E8F0; vertical-align:top; }}
                    .data-table tr:nth-child(even) {{ background:#F8FAFC; }}
                    .chart-card.fullscreen {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 9999; margin: 0; border-radius: 0; padding: 50px 20px 20px 20px; box-sizing: border-box; }}
                    .factions-container {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; }}
                    .faction-box {{ background: #F9F9F9; border: 1px solid #E0E0E0; padding: 12px 15px; border-radius: 4px; }}
                    .faction-box h4 {{ margin: 0 0 8px 0; color: #333; font-size: 15px; }}
                    .faction-box p {{ margin: 0; font-size: 13px; color: #555; line-height: 1.5; }}
                    .info-title-container {{ position: absolute; top: 15px; left: 15px; z-index: 50; pointer-events: none; }}
                    .custom-tooltip {{ pointer-events: auto; position: relative; display: inline-block; cursor: help; color: #005A9E; font-size: 16px; margin-left: 8px; }}
                    .custom-tooltip .custom-tooltip-text {{ visibility: hidden; width: 320px; background-color: #333; color: #fff; text-align: left; border-radius: 6px; padding: 15px; font-size: 13px; position: absolute; z-index: 1000; bottom: 125%; left: -10px; opacity: 0; transition: opacity 0.3s; box-shadow: 0 4px 8px rgba(0,0,0,0.2); line-height: 1.4; }}
                    .custom-tooltip:hover .custom-tooltip-text {{ visibility: visible; opacity: 1; }}
                    .export-data-link {{ display:inline-block; padding:7px 11px; border:1px solid #99C9FF; border-radius:5px; background:#E1F0FF; color:#005A9E; text-decoration:none; font-weight:bold; font-size:12px; }}
                    .interactive-filter-bar {{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin:-15px auto 22px; max-width:1200px; padding:10px 12px; background:#FFFFFF; border:1px solid #E2E8F0; border-radius:7px; box-shadow:0 2px 4px rgba(0,0,0,.04); }}
                    .interactive-filter-label {{ font-size:12px; font-weight:bold; color:#475569; }}
                    .filter-chip {{ display:inline-flex; align-items:center; gap:5px; padding:4px 8px; border-radius:14px; background:#E1F0FF; border:1px solid #99C9FF; color:#005A9E; font-size:11px; font-weight:bold; }}
                    .filter-chip button {{ border:0; background:transparent; cursor:pointer; color:#005A9E; font-weight:bold; padding:0; }}
                    .clear-crossfilters {{ border:1px solid #CBD5E1; background:#F8FAFC; color:#475569; border-radius:4px; padding:4px 8px; cursor:pointer; font-size:11px; }}
                    .crossfilter-summary {{ margin-left:auto; color:#64748B; font-size:11px; }}
                    .data-tools {{ display:flex; gap:7px; align-items:center; margin:0 0 8px 0; flex-wrap:wrap; }}
                    .data-search {{ flex:1 1 220px; min-width:160px; border:1px solid #CBD5E1; border-radius:4px; padding:6px 8px; font-size:12px; }}
                    .data-tools button {{ cursor:pointer; border:1px solid #CBD5E1; background:#F8FAFC; color:#334155; border-radius:4px; padding:5px 8px; font-size:11px; font-weight:bold; }}
                    .data-row-count {{ color:#64748B; font-size:11px; white-space:nowrap; }}
                    .data-table th {{ cursor:pointer; user-select:none; }}
                    .sort-indicator {{ color:#94A3B8; font-size:10px; }}
                    .data-table tr.table-filter-hidden {{ display:none; }}
                    .plotly-click-hint {{ text-align:center; color:#64748B; font-size:11px; margin-top:-4px; }}
                </style>
                <script>
                    function switchAI(selectedId) {{
                        const sections = document.querySelectorAll('.dashboard-view-panel');
                        sections.forEach(sec => {{ sec.style.display = 'none'; }});

                        const activeSec = document.getElementById(selectedId);
                        if(activeSec) {{
                            activeSec.style.display = 'block';
                            window.dispatchEvent(new Event('resize'));
                            setTimeout(() => {{ wirePlotlyCrossFiltering(); renderDashboardFilters(); }}, 40);
                        }}
                    }}
                    function toggleChartData(btn, mode) {{
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
                    }}

                    function toggleFullscreen(btn) {{
                        const card = btn.parentElement;
                        card.classList.toggle('fullscreen');
                        btn.innerHTML = card.classList.contains('fullscreen') ? '✖ Close' : '⛶ Expand';
                        setTimeout(() => {{ window.dispatchEvent(new Event('resize')); }}, 50);
                    }}

                    const dashboardFilters = {{company: new Set(), ai: new Set(), status: new Set()}};
                    let dashboardRows = [];

                    function currentScopeRows() {{
                        const selected = document.getElementById('scope-select')?.value || 'global';
                        let rows = dashboardRows;
                        if(selected.startsWith('wi_')) {{
                            const option = document.querySelector(`#scope-select option[value="${{selected}}"]`);
                            const wi = option?.dataset?.wi || '';
                            if(wi) rows = rows.filter(r => (r.wis || []).includes(wi));
                        }} else if(selected.startsWith('ai_')) {{
                            const option = document.querySelector(`#scope-select option[value="${{selected}}"]`);
                            const ai = option?.dataset?.ai || '';
                            if(ai) rows = rows.filter(r => r.ai === ai);
                        }}
                        return rows;
                    }}

                    function filteredRows() {{
                        return currentScopeRows().filter(r =>
                            (!dashboardFilters.company.size || (r.companies || []).some(c => dashboardFilters.company.has(c))) &&
                            (!dashboardFilters.ai.size || dashboardFilters.ai.has(r.ai)) &&
                            (!dashboardFilters.status.size || dashboardFilters.status.has(r.status))
                        );
                    }}

                    function addDashboardFilter(dimension, value) {{
                        if(!value || !dashboardFilters[dimension]) return;
                        dashboardFilters[dimension].add(String(value));
                        renderDashboardFilters();
                    }}
                    function removeDashboardFilter(dimension, value) {{
                        dashboardFilters[dimension]?.delete(String(value));
                        renderDashboardFilters();
                    }}
                    function clearDashboardFilters() {{
                        Object.values(dashboardFilters).forEach(s => s.clear());
                        renderDashboardFilters();
                    }}
                    function renderDashboardFilters() {{
                        const host = document.getElementById('active-crossfilters');
                        if(!host) return;
                        host.innerHTML = '';
                        const names = {{company:'Company', ai:'AI', status:'Status'}};
                        Object.entries(dashboardFilters).forEach(([dim, values]) => {{
                            values.forEach(value => {{
                                const chip = document.createElement('span');
                                chip.className='filter-chip';
                                chip.append(document.createTextNode(`${{names[dim]}}: ${{value}}`));
                                const b=document.createElement('button'); b.textContent='×';
                                b.onclick=()=>removeDashboardFilter(dim,value); chip.appendChild(b);
                                host.appendChild(chip);
                            }});
                        }});
                        const rows=filteredRows();
                        const summary=document.getElementById('crossfilter-summary');
                        if(summary) summary.textContent=`${{rows.length}} TDocs match active scope + filters`;
                        applyTableCrossFilters();
                    }}

                    function tableCardDimension(card) {{
                        const txt=(card?.querySelector('.custom-tooltip-text')?.textContent || '').toLowerCase();
                        if(txt.includes('top contributors') || txt.includes('specialization')) return 'company';
                        if(txt.includes('agenda items by volume') || txt.includes('agenda items by status') || txt.includes('revision activity')) return 'ai';
                        if(txt.includes('tdoc outcomes')) return 'status';
                        return '';
                    }}
                    function applyTableCrossFilters() {{
                        document.querySelectorAll('.dashboard-view-panel').forEach(panel => {{
                            if(panel.style.display==='none') return;
                            panel.querySelectorAll('.data-table').forEach(table => {{
                                const card=table.closest('.chart-card');
                                const dim=tableCardDimension(card);
                                const allowed=filteredRows();
                                let values=null;
                                if(dim==='company') values=new Set(allowed.flatMap(r=>r.companies||[]));
                                if(dim==='ai') values=new Set(allowed.map(r=>r.ai));
                                if(dim==='status') values=new Set(allowed.map(r=>r.status));
                                Array.from(table.tBodies[0]?.rows || []).forEach(row => {{
                                    if(!values || !dim) return;
                                    const value=(row.cells[0]?.dataset.value || row.cells[0]?.textContent || '').trim();
                                    row.dataset.crossVisible = values.has(value) ? '1':'0';
                                }});
                                const search=table.closest('.data-pane')?.querySelector('.data-search');
                                if(search) filterStatsTable(search); else refreshTableVisibility(table,'');
                            }});
                        }});
                    }}

                    function refreshTableVisibility(table, query) {{
                        query=(query||'').toLowerCase();
                        let visible=0, total=0;
                        Array.from(table.tBodies[0]?.rows || []).forEach(row => {{
                            total++;
                            const searchOk=!query || row.textContent.toLowerCase().includes(query);
                            const crossOk=row.dataset.crossVisible !== '0';
                            row.classList.toggle('table-filter-hidden', !(searchOk && crossOk));
                            if(searchOk && crossOk) visible++;
                        }});
                        const count=table.closest('.data-pane')?.querySelector('.data-row-count');
                        if(count) count.textContent=`${{visible}} / ${{total}} rows`;
                    }}
                    function filterStatsTable(input) {{
                        const table=input.closest('.data-pane')?.querySelector('.data-table');
                        if(table) refreshTableVisibility(table,input.value);
                    }}
                    function parseSortValue(v) {{
                        const cleaned=String(v).trim().replace(/[% ,]/g,'');
                        return cleaned!=='' && !Number.isNaN(Number(cleaned)) ? Number(cleaned) : String(v).toLowerCase();
                    }}
                    function sortStatsTable(th) {{
                        const table=th.closest('table'), body=table?.tBodies[0]; if(!body) return;
                        const idx=Array.from(th.parentNode.children).indexOf(th);
                        const asc=th.dataset.sort !== 'asc';
                        table.querySelectorAll('th').forEach(h=>{{h.dataset.sort=''; const i=h.querySelector('.sort-indicator'); if(i)i.textContent='↕';}});
                        th.dataset.sort=asc?'asc':'desc'; const indicator=th.querySelector('.sort-indicator'); if(indicator)indicator.textContent=asc?'↑':'↓';
                        const rows=Array.from(body.rows);
                        rows.sort((a,b)=>{{
                            const av=parseSortValue(a.cells[idx]?.dataset.value ?? a.cells[idx]?.textContent ?? '');
                            const bv=parseSortValue(b.cells[idx]?.dataset.value ?? b.cells[idx]?.textContent ?? '');
                            const cmp=(typeof av==='number' && typeof bv==='number') ? av-bv : String(av).localeCompare(String(bv),undefined,{{numeric:true,sensitivity:'base'}});
                            return asc?cmp:-cmp;
                        }});
                        rows.forEach(r=>body.appendChild(r));
                    }}
                    function visibleTableMatrix(btn) {{
                        const pane=btn.closest('.data-pane'), table=pane?.querySelector('.data-table'); if(!table)return [];
                        const headers=Array.from(table.tHead?.rows[0]?.cells||[]).map(c=>c.textContent.replace(/[↕↑↓]/g,'').trim());
                        const rows=Array.from(table.tBodies[0]?.rows||[]).filter(r=>!r.classList.contains('table-filter-hidden'))
                            .map(r=>Array.from(r.cells).map(c=>c.dataset.value ?? c.textContent.trim()));
                        return [headers,...rows];
                    }}
                    function csvEscape(v){{const s=String(v??''); return /[",\n]/.test(s)?`"${{s.replace(/"/g,'""')}}"`:s;}}
                    function copyVisibleTable(btn) {{
                        const m=visibleTableMatrix(btn); if(!m.length)return;
                        navigator.clipboard?.writeText(m.map(r=>r.join('\t')).join('\n'));
                        const old=btn.textContent; btn.textContent='✓ Copied'; setTimeout(()=>btn.textContent=old,900);
                    }}
                    function downloadVisibleTableCsv(btn) {{
                        const m=visibleTableMatrix(btn); if(!m.length)return;
                        const csv='\uFEFF'+m.map(r=>r.map(csvEscape).join(',')).join('\r\n');
                        const blob=new Blob([csv],{{type:'text/csv;charset=utf-8;'}}), url=URL.createObjectURL(blob);
                        const a=document.createElement('a'); a.href=url; a.download='statistics_table.csv'; a.click(); URL.revokeObjectURL(url);
                    }}
                    function wirePlotlyCrossFiltering() {{
                        document.querySelectorAll('.js-plotly-plot').forEach(plot => {{
                            if(plot.dataset.crossfilterWired) return;
                            plot.dataset.crossfilterWired='1';
                            plot.on('plotly_click', ev => {{
                                const pt=ev?.points?.[0]; if(!pt) return;
                                const dim=pt.data?.meta?.filterDimension; if(!dim) return;
                                let value='';
                                if(pt.customdata && Array.isArray(pt.customdata)) value=pt.customdata[0];
                                else if(pt.customdata != null) value=pt.customdata;
                                else if(dim==='company') value=pt.y;
                                else if(dim==='status') value=pt.label;
                                if(value) addDashboardFilter(dim,String(value));
                            }});
                        }});
                    }}
                    document.addEventListener('DOMContentLoaded', () => {{
                        const raw=document.getElementById('dashboard-filter-data');
                        try {{ dashboardRows=JSON.parse(raw?.textContent || '[]'); }} catch(e) {{ dashboardRows=[]; }}
                        wirePlotlyCrossFiltering();
                        document.querySelectorAll('.data-search').forEach(filterStatsTable);
                        renderDashboardFilters();
                    }});
                </script>
            </head>
            <body>
                <h1>📊 {meeting_name} - TDoc Analytics Dashboard</h1>

                <div class="selector-container">
                    <label for="scope-select">🎯 Analytics Scope / Agenda Item:</label>
                    <select id="scope-select" onchange="switchAI(this.value)">
                        {" ".join(dropdown_options)}
                    </select>
                </div>
                <div class="interactive-filter-bar">
                    <span class="interactive-filter-label">🔎 Interactive filters:</span>
                    <span id="active-crossfilters"></span>
                    <button type="button" class="clear-crossfilters" onclick="clearDashboardFilters()">Clear filters</button>
                    <span id="crossfilter-summary" class="crossfilter-summary"></span>
                </div>
                <script type="application/json" id="dashboard-filter-data">{interactive_json}</script>

                {" ".join(views_html_buffer)}
            </body>
            </html>
            """

            xlsx_path = self.export_dir / "Statistics_Data.xlsx"
            self._export_statistics_workbook(df, xlsx_path)
            dashboard_template = dashboard_template.replace(
                '<h1>📊',
                '<div style="text-align:right;margin-bottom:8px;"><a class="export-data-link" href="Statistics_Data.xlsx">📥 Statistics Data (Excel)</a></div><h1>📊',
                1,
            )
            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_template)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))

    def _export_statistics_workbook(self, df: pd.DataFrame, path: Path):
        """Write reusable aggregated statistics behind the HTML dashboard."""
        ai_cols = ['Agenda Item', 'AI_Acronym', 'AI_Topic']
        ai_volume = (df.groupby(ai_cols, dropna=False).size().reset_index(name='TDocs')
                     .sort_values('TDocs', ascending=False))
        status = df['TDoc Status'].fillna('').astype(str).value_counts().rename_axis('Status').reset_index(name='TDocs')
        company_rows = []
        for _, row in df.iterrows():
            for company in set(row.get('Clean_Companies', []) or []):
                company_rows.append({'Company': company, 'Agenda Item': row.get('Agenda Item', ''),
                                     'WI': '; '.join(row.get('Clean_WI_List', []) or []), 'TDoc': row.get('TDoc', '')})
        company_df = pd.DataFrame(company_rows)
        company_volume = (company_df.groupby('Company').size().reset_index(name='TDocs').sort_values('TDocs', ascending=False)
                          if not company_df.empty else pd.DataFrame(columns=['Company', 'TDocs']))
        wi_rows = []
        for _, row in df.iterrows():
            for wi in row.get('Clean_WI_List', []) or []:
                wi_rows.append({'WI/SI': wi, 'TDoc': row.get('TDoc', ''), 'Agenda Item': row.get('Agenda Item', ''),
                                'Status': row.get('TDoc Status', '')})
        wi_df = pd.DataFrame(wi_rows)
        wi_volume = (wi_df.groupby('WI/SI').size().reset_index(name='TDocs').sort_values('TDocs', ascending=False)
                     if not wi_df.empty else pd.DataFrame(columns=['WI/SI', 'TDocs']))
        revision = build_revision_stats(df).drop(columns=['_Axis'], errors='ignore')
        specialization = build_company_specialization(df)
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            ai_volume.to_excel(writer, sheet_name='AI Volume', index=False)
            status.to_excel(writer, sheet_name='Outcomes', index=False)
            wi_volume.to_excel(writer, sheet_name='WI-SI Volume', index=False)
            company_volume.to_excel(writer, sheet_name='Company Volume', index=False)
            revision.to_excel(writer, sheet_name='Revision Activity', index=False)
            specialization.to_excel(writer, sheet_name='Specialization', index=False)
            if not company_df.empty:
                pd.crosstab(company_df['Company'], company_df['Agenda Item']).to_excel(writer, sheet_name='Company x AI')
            if not wi_df.empty:
                pd.crosstab(wi_df['WI/SI'], wi_df['Agenda Item']).to_excel(writer, sheet_name='WI-SI x AI')

    def _compile_view_block(self, scope_id, total_tdocs, total_companies, ai_volume_html, status_html, comp_html,
                            net_html, cluster_html, cohesion_html, list_html,
                            ai_status_html=None, heatmap_html=None, ai_volume_table="", status_table="",
                            comp_table="", ai_status_table="", heatmap_table="", network_table="", cluster_table="",
                            cohesion_table="", revision_html="", revision_table="", specialization_html="", specialization_table="", is_visible=False):

        display_style = "block" if is_visible else "none"

        volume_card = ""
        if ai_volume_html:
            volume_card = """
            <div class="chart-card">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Agenda Items by Volume:</b> Displays top topics based on submitted document frequency.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__AI_VOLUME_HTML__</div><div class="data-pane">__AI_VOLUME_TABLE__</div>
            </div>
            """.replace("__AI_VOLUME_HTML__", str(ai_volume_html)).replace("__AI_VOLUME_TABLE__", str(ai_volume_table))

        ai_status_card = ""
        if ai_status_html:
            ai_status_card = """
            <div class="chart-card" style="grid-column: 1 / -1; height: 500px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Agenda Items by Status:</b> Breakdown of outcomes across the top topics.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__AI_STATUS_HTML__</div><div class="data-pane">__AI_STATUS_TABLE__</div>
            </div>
            """.replace("__AI_STATUS_HTML__", str(ai_status_html)).replace("__AI_STATUS_TABLE__", str(ai_status_table))

        heatmap_card = ""
        if heatmap_html:
            heatmap_card = """
            <div class="chart-card" style="grid-column: 1 / -1; height: 650px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Company Focus Matrix:</b> Heatmap of TDoc submissions by the top companies across top topics.</span></span>
                </div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__HEATMAP_HTML__</div><div class="data-pane">__HEATMAP_TABLE__</div>
            </div>
            """.replace("__HEATMAP_HTML__", str(heatmap_html)).replace("__HEATMAP_TABLE__", str(heatmap_table))

        revision_card = ""
        if revision_html:
            revision_card = """
            <div class="chart-card" style="grid-column: 1 / -1; height: 520px;">
                <div class="info-title-container"><span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Revision Activity:</b> Shows original TDocs and subsequent revisions per Agenda Item. Revision intensity is descriptive and does not imply quality.</span></span></div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__REVISION_HTML__</div><div class="data-pane">__REVISION_TABLE__</div>
            </div>
            """.replace("__REVISION_HTML__", str(revision_html)).replace("__REVISION_TABLE__", str(revision_table))

        specialization_card = ""
        if specialization_html:
            specialization_card = """
            <div class="chart-card" style="grid-column: 1 / -1; height: 560px;">
                <div class="info-title-container"><span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Company Topic Specialization:</b> Bubble size is contribution volume; horizontal position is the share of each company's TDocs concentrated in its most frequent topic.</span></span></div>
                <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                <div class="chart-pane">__SPECIALIZATION_HTML__</div><div class="data-pane">__SPECIALIZATION_TABLE__</div>
            </div>
            """.replace("__SPECIALIZATION_HTML__", str(specialization_html)).replace("__SPECIALIZATION_TABLE__", str(specialization_table))

        grid_col_span = "" if scope_id != "global" else "grid-column: span 1;"

        html_template = """
        <div id="__SCOPE_ID__" class="dashboard-view-panel" style="display: __DISPLAY_STYLE__;">
            <div class="kpi-container">
                <div class="kpi-card"><h3>__TOTAL_TDOCS__</h3><p>View TDocs</p></div>
                <div class="kpi-card"><h3>__TOTAL_COMPANIES__</h3><p>Active Companies</p></div>
            </div>

            <div class="grid-container">
                __VOLUME_CARD__
                <div class="chart-card" style="__GRID_COL_SPAN__">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>TDoc Outcomes:</b> Breakdown of actions applied across this data selection.</span></span>
                    </div>
                    <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    <div class="chart-pane">__STATUS_HTML__</div><div class="data-pane">__STATUS_TABLE__</div>
                </div>

                __AI_STATUS_CARD__
                __HEATMAP_CARD__

                <div class="chart-card" style="grid-column: 1 / -1; height: 550px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Top Contributors:</b> Active entities in the scope subset.</span></span>
                    </div>
                    <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    <div class="chart-pane">__COMP_HTML__</div><div class="data-pane">__COMP_TABLE__</div>
                </div>

                <div class="chart-card" style="grid-column: 1 / -1; height: 750px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Strategic Alliances:</b> Subset collaboration mappings under fixed global configurations.</span></span>
                    </div>
                    <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    <div class="chart-pane">__NET_HTML__</div><div class="data-pane">__NETWORK_TABLE__</div>
                </div>

                <div class="chart-card" style="height: 450px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Faction Output Volume:</b> Quantifies documents matching this view scope authored by the tracked alliance clusters.</span></span>
                    </div>
                    <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    <div class="chart-pane">__CLUSTER_HTML__</div><div class="data-pane">__CLUSTER_TABLE__</div>
                </div>

                <div class="chart-card" style="height: 450px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Cohesion Tracker:</b> Displays the interaction density of factions under this specific subset.</span></span>
                    </div>
                    <div class="view-toggle"><button class="view-btn active" onclick="toggleChartData(this,'chart')">Chart</button><button class="view-btn" onclick="toggleChartData(this,'data')">Data</button></div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    <div class="chart-pane">__COHESION_HTML__</div><div class="data-pane">__COHESION_TABLE__</div>
                </div>

                __REVISION_CARD__
                __SPECIALIZATION_CARD__

                <div class="chart-card" style="grid-column: 1 / -1; height: auto; padding: 20px;">
                    __LIST_HTML__
                </div>
            </div>
        </div>
        """

        html_template = html_template.replace("__SCOPE_ID__", str(scope_id))
        html_template = html_template.replace("__DISPLAY_STYLE__", str(display_style))
        html_template = html_template.replace("__TOTAL_TDOCS__", str(total_tdocs))
        html_template = html_template.replace("__TOTAL_COMPANIES__", str(total_companies))
        html_template = html_template.replace("__VOLUME_CARD__", str(volume_card))
        html_template = html_template.replace("__GRID_COL_SPAN__", str(grid_col_span))
        html_template = html_template.replace("__STATUS_HTML__", str(status_html))
        html_template = html_template.replace("__STATUS_TABLE__", str(status_table))

        # Inject the new blocks
        html_template = html_template.replace("__AI_STATUS_CARD__", str(ai_status_card))
        html_template = html_template.replace("__HEATMAP_CARD__", str(heatmap_card))

        html_template = html_template.replace("__COMP_HTML__", str(comp_html))
        html_template = html_template.replace("__COMP_TABLE__", str(comp_table))
        html_template = html_template.replace("__NET_HTML__", str(net_html))
        html_template = html_template.replace("__NETWORK_TABLE__", str(network_table))
        html_template = html_template.replace("__CLUSTER_HTML__", str(cluster_html))
        html_template = html_template.replace("__CLUSTER_TABLE__", str(cluster_table))
        html_template = html_template.replace("__COHESION_HTML__", str(cohesion_html))
        html_template = html_template.replace("__COHESION_TABLE__", str(cohesion_table))
        html_template = html_template.replace("__REVISION_CARD__", str(revision_card))
        html_template = html_template.replace("__SPECIALIZATION_CARD__", str(specialization_card))
        html_template = html_template.replace("__LIST_HTML__", str(list_html))

        return html_template