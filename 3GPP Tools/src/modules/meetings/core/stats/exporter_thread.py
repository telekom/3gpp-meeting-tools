# --- File: src/modules/meetings/core/stats/exporter_thread.py ---
import os
import re
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
import plotly.express as px

from core.utils.company_sanitizer import CompanySanitizer
from .plot_agenda import generate_ai_volume_plot
from .plot_status import generate_outcomes_plot
from .plot_contributors import generate_top_contributors_plot
from .plot_alliances import compute_global_communities, generate_alliance_plots


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
        self.cfg_export_html = self.config.get("export_html_plots", False)  # ---> Retrieve the new boolean toggle

        self.THEME_COLOR = '#005A9E'
        self.PALETTE = px.colors.qualitative.Plotly
        self.CLUSTER_PALETTE = px.colors.qualitative.Alphabet

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

            df = df[~df['TDoc Status'].str.lower().str.contains('withdrawn', na=False)].copy()
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            global_factions = compute_global_communities(df, self.cfg_resolution)

            # Pass the toggle into the global layout blocks
            g_html_ai = generate_ai_volume_plot(df, plots_dir, self.THEME_COLOR, prefix_id="Global",
                                                save_html=self.cfg_export_html)
            g_html_status = generate_outcomes_plot(df, plots_dir, self.PALETTE, prefix_id="Global",
                                                   save_html=self.cfg_export_html)
            g_html_comp, total_companies = generate_top_contributors_plot(df, plots_dir, self.THEME_COLOR,
                                                                          self.cfg_top_count, prefix_id="Global",
                                                                          save_html=self.cfg_export_html)
            g_html_net, g_html_cluster, g_html_cohesion, g_html_list = generate_alliance_plots(df, plots_dir,
                                                                                               self.cfg_threshold,
                                                                                               self.CLUSTER_PALETTE,
                                                                                               global_factions,
                                                                                               prefix_id="Global",
                                                                                               save_html=self.cfg_export_html)

            raw_ais = df['Agenda Item'].dropna().unique()

            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

            unique_ais = sorted([str(ai).strip() for ai in raw_ais if str(ai).strip()], key=natural_sort_key)

            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"

            views_html_buffer = []
            dropdown_options = ['<option value="global">🌐 Overall Meeting View</option>']

            views_html_buffer.append(
                self._compile_view_block("global", len(df), total_companies, g_html_ai, g_html_status, g_html_comp,
                                         g_html_net, g_html_cluster, g_html_cohesion, g_html_list, is_visible=True))

            for idx, ai_name in enumerate(unique_ais):
                ai_df = df[df['Agenda Item'].str.strip() == ai_name].copy()
                if ai_df.empty: continue

                safe_id = f"ai_{idx}"
                clean_ai_name = re.sub(r'[\\/*?:\"<>|]', '_', str(ai_name))
                safe_ai_prefix = "AI_" + clean_ai_name

                dropdown_options.append(
                    f'<option value="{safe_id}">📌 Agenda Item {ai_name} ({len(ai_df)} TDocs)</option>')

                # Pass the toggle into the iterative blocks
                ai_html_status = generate_outcomes_plot(ai_df, plots_dir, self.PALETTE, safe_ai_prefix,
                                                        save_html=self.cfg_export_html)
                ai_html_comp, ai_companies = generate_top_contributors_plot(ai_df, plots_dir, self.THEME_COLOR,
                                                                            self.cfg_top_count, safe_ai_prefix,
                                                                            save_html=self.cfg_export_html)
                ai_html_net, ai_html_cluster, ai_html_cohesion, ai_html_list = generate_alliance_plots(ai_df, plots_dir,
                                                                                                       self.cfg_threshold,
                                                                                                       self.CLUSTER_PALETTE,
                                                                                                       global_factions,
                                                                                                       safe_ai_prefix,
                                                                                                       save_html=self.cfg_export_html)

                views_html_buffer.append(self._compile_view_block(
                    safe_id, len(ai_df), ai_companies,
                    ai_volume_html=None, status_html=ai_html_status, comp_html=ai_html_comp,
                    net_html=ai_html_net, cluster_html=ai_html_cluster, cohesion_html=ai_html_cohesion,
                    list_html=ai_html_list,
                    is_visible=False
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
                    .chart-card > div {{ flex-grow: 1; width: 100%; height: 100%; }}
                    .fs-btn {{ position: absolute; top: 10px; right: 10px; z-index: 100; cursor: pointer; background: #E1F0FF; color: #005A9E; border: 1px solid #99C9FF; border-radius: 4px; padding: 5px 10px; font-weight: bold; font-size: 12px; }}
                    .fs-btn:hover {{ background: #CCE4FF; }}
                    .chart-card.fullscreen {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 9999; margin: 0; border-radius: 0; padding: 50px 20px 20px 20px; box-sizing: border-box; }}
                    .factions-container {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; }}
                    .faction-box {{ background: #F9F9F9; border: 1px solid #E0E0E0; padding: 12px 15px; border-radius: 4px; }}
                    .faction-box h4 {{ margin: 0 0 8px 0; color: #333; font-size: 15px; }}
                    .faction-box p {{ margin: 0; font-size: 13px; color: #555; line-height: 1.5; }}
                    .info-title-container {{ position: absolute; top: 15px; left: 15px; z-index: 50; pointer-events: none; }}
                    .custom-tooltip {{ pointer-events: auto; position: relative; display: inline-block; cursor: help; color: #005A9E; font-size: 16px; margin-left: 8px; }}
                    .custom-tooltip .custom-tooltip-text {{ visibility: hidden; width: 320px; background-color: #333; color: #fff; text-align: left; border-radius: 6px; padding: 15px; font-size: 13px; position: absolute; z-index: 1000; bottom: 125%; left: -10px; opacity: 0; transition: opacity 0.3s; box-shadow: 0 4px 8px rgba(0,0,0,0.2); line-height: 1.4; }}
                    .custom-tooltip:hover .custom-tooltip-text {{ visibility: visible; opacity: 1; }}
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

            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_template)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))

    def _compile_view_block(self, scope_id, total_tdocs, total_companies, ai_volume_html, status_html, comp_html,
                            net_html, cluster_html, cohesion_html, list_html, is_visible=False):
        # (This method remains exactly the same as your current functioning version)
        display_style = "block" if is_visible else "none"

        volume_card = ""
        if ai_volume_html:
            volume_card = """
            <div class="chart-card">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Agenda Items by Volume:</b> Displays top topics based on submitted document frequency.</span></span>
                </div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                __AI_VOLUME_HTML__
            </div>
            """.replace("__AI_VOLUME_HTML__", str(ai_volume_html))

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
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    __STATUS_HTML__
                </div>

                <div class="chart-card" style="grid-column: 1 / -1; height: 550px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Top Contributors:</b> Active entities in the scope subset.</span></span>
                    </div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    __COMP_HTML__
                </div>

                <div class="chart-card" style="grid-column: 1 / -1; height: 750px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Strategic Alliances:</b> Subset collaboration mappings under fixed global configurations.</span></span>
                    </div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    __NET_HTML__
                </div>

                <div class="chart-card" style="height: 450px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Faction Output Volume:</b> Quantifies documents matching this view scope authored by the tracked alliance clusters.</span></span>
                    </div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    __CLUSTER_HTML__
                </div>

                <div class="chart-card" style="height: 450px;">
                    <div class="info-title-container">
                        <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Cohesion Tracker:</b> Displays the interaction density of factions under this specific subset.</span></span>
                    </div>
                    <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                    __COHESION_HTML__
                </div>

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
        html_template = html_template.replace("__COMP_HTML__", str(comp_html))
        html_template = html_template.replace("__NET_HTML__", str(net_html))
        html_template = html_template.replace("__CLUSTER_HTML__", str(cluster_html))
        html_template = html_template.replace("__COHESION_HTML__", str(cohesion_html))
        html_template = html_template.replace("__LIST_HTML__", str(list_html))

        return html_template