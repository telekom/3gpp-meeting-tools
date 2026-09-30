# --- File: src/modules/meetings/core/stats/exporter_thread.py ---
import os
import re
import html
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
                    f'<option value="{safe_id}">🧩 {html.escape(wi_name)} ({len(wi_df)} TDocs)</option>')
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
                    f'<option value="{safe_id}">📌 {html.escape(ai_display)} ({len(ai_df)} TDocs)</option>')

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
                </style>
                <script>
                    function switchAI(selectedId) {{
                        const sections = document.querySelectorAll('.dashboard-view-panel');
                        sections.forEach(sec => {{ sec.style.display = 'none'; }});

                        const activeSec = document.getElementById(selectedId);
                        if(activeSec) {{
                            activeSec.style.display = 'block';
                            window.dispatchEvent(new Event('resize'));
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