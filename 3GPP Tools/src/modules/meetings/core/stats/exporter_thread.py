# --- File: src/modules/meetings/core/stats/exporter_thread.py ---
import os
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
import plotly.express as px

from core.utils.company_sanitizer import CompanySanitizer
from .plot_agenda import generate_ai_volume_plot
from .plot_status import generate_outcomes_plot
from .plot_contributors import generate_top_contributors_plot
from .plot_alliances import generate_alliance_plots


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

        self.THEME_COLOR = '#005A9E'
        self.PALETTE = px.colors.qualitative.Plotly
        self.CLUSTER_PALETTE = px.colors.qualitative.Alphabet

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            df = pd.DataFrame(self.tdocs_data)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            # Data Sanitization
            df = df[~df['TDoc Status'].str.lower().str.contains('withdrawn', na=False)].copy()
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            # --- Generate Modular Plots ---
            html_ai = generate_ai_volume_plot(df, self.export_dir, self.THEME_COLOR)
            html_status = generate_outcomes_plot(df, self.export_dir, self.PALETTE)
            html_comp, total_companies = generate_top_contributors_plot(df, self.export_dir, self.THEME_COLOR,
                                                                        self.cfg_top_count)

            html_net, html_cluster_contribs, html_cohesion_plot, html_faction_list = generate_alliance_plots(
                df, self.export_dir, self.cfg_threshold, self.cfg_resolution, self.CLUSTER_PALETTE
            )

            # --- Compile Dashboard HTML ---
            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"
            total_tdocs = len(df)

            dashboard_html = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>3GPP Statistics - __MEETING_NAME__</title>
                <style>
                    body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }
                    h1 { color: #333; text-align: center; margin-bottom: 30px; }
                    .kpi-container { display: flex; justify-content: center; gap: 20px; margin-bottom: 40px; }
                    .kpi-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 200px; border-top: 4px solid #005A9E; }
                    .kpi-card h3 { margin: 0; font-size: 32px; color: #005A9E; }
                    .kpi-card p { margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }
                    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }

                    .chart-card { 
                        position: relative; background: white; border-radius: 8px; 
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 40px 15px 15px 15px; 
                        height: 500px; display: flex; flex-direction: column; transition: all 0.3s ease; 
                    }
                    .chart-card > div { flex-grow: 1; width: 100%; height: 100%; }

                    .fs-btn { position: absolute; top: 10px; right: 10px; z-index: 100; cursor: pointer; background: #E1F0FF; color: #005A9E; border: 1px solid #99C9FF; border-radius: 4px; padding: 5px 10px; font-weight: bold; font-size: 12px; transition: background 0.2s; }
                    .fs-btn:hover { background: #CCE4FF; }

                    .chart-card.fullscreen { 
                        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; 
                        z-index: 9999; margin: 0; border-radius: 0; 
                        padding: 50px 20px 20px 20px; box-sizing: border-box; 
                    }

                    .factions-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; }
                    .faction-box { background: #F9F9F9; border: 1px solid #E0E0E0; padding: 12px 15px; border-radius: 4px; }
                    .faction-box h4 { margin: 0 0 8px 0; color: #333; font-size: 15px; }
                    .faction-box p { margin: 0; font-size: 13px; color: #555; line-height: 1.5; }

                    .info-title-container { position: absolute; top: 15px; left: 15px; z-index: 50; }
                    .tooltip { position: relative; display: inline-block; cursor: help; color: #005A9E; font-size: 16px; margin-left: 8px; }
                    .tooltip .tooltip-text {
                        visibility: hidden; width: 320px; background-color: #333; color: #fff; 
                        text-align: left; border-radius: 6px; padding: 15px; font-size: 13px; font-weight: normal;
                        position: absolute; z-index: 1000; bottom: 125%; left: -10px; 
                        opacity: 0; transition: opacity 0.3s; box-shadow: 0 4px 8px rgba(0,0,0,0.2); line-height: 1.4;
                    }
                    .tooltip .tooltip-text::after {
                        content: ""; position: absolute; top: 100%; left: 15px; 
                        border-width: 5px; border-style: solid; border-color: #333 transparent transparent transparent;
                    }
                    .tooltip:hover .tooltip-text { visibility: visible; opacity: 1; }
                </style>
                <script>
                    function toggleFullscreen(btn) {
                        const card = btn.parentElement;
                        card.classList.toggle('fullscreen');
                        if (card.classList.contains('fullscreen')) {
                            btn.innerHTML = '✖ Close';
                        } else {
                            btn.innerHTML = '⛶ Expand';
                        }
                        setTimeout(() => { window.dispatchEvent(new Event('resize')); }, 50);
                    }
                </script>
            </head>
            <body>
                <h1>📊 __MEETING_NAME__ - TDoc Statistics Dashboard</h1>

                <div class="kpi-container">
                    <div class="kpi-card"><h3>__TOTAL_TDOCS__</h3><p>Total TDocs</p></div>
                    <div class="kpi-card"><h3>__TOTAL_COMPANIES__</h3><p>Participating Companies</p></div>
                </div>

                <div class="grid-container">
                    <div class="chart-card">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>Agenda Items by Volume:</b> Displays the top topics based on the sheer number of submitted documents. This helps identify where the majority of the working group's effort and debate is currently focused.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_AI__
                    </div>

                    <div class="chart-card">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>TDoc Outcomes:</b> A breakdown of the final decisions made on the submitted documents (e.g., Agreed, Revised, Noted). Note that 'Withdrawn' documents are explicitly excluded from this dataset.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_STATUS__
                    </div>

                    <div class="chart-card" style="grid-column: 1 / -1; height: 600px;">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>Top Contributors:</b> Ranks the most active companies based on the total number of documents they have either authored or co-signed in this meeting.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_COMP__
                    </div>

                    <div class="chart-card" style="grid-column: 1 / -1; height: 750px;">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>Strategic Alliances:</b> Visualizes the collaboration network. Each node represents a company. The lines connecting them represent co-signed documents. Thicker lines indicate a stronger alliance with a higher volume of shared documents.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_NET__
                    </div>

                    <div class="chart-card" style="height: 450px;">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>Louvain Method:</b> A mathematical algorithm that automatically discovers distinct "communities" or "factions" within a network by finding groups of companies that co-sign with each other significantly more often than they co-sign with outsiders.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_CLUSTER_CONTRIBS__
                    </div>

                    <div class="chart-card" style="height: 450px;">
                        <div class="info-title-container">
                            <span class="tooltip">ⓘ
                                <span class="tooltip-text">
                                    <b>Cohesion Score (Network Density):</b> Measures how tightly-knit a faction is on a scale of 0 to 1.<br><br>It is calculated by dividing the actual number of co-signs within the faction by the maximum possible number of co-signs (if every single member explicitly partnered with every other member). A higher score means strong, uniform coordination.
                                </span>
                            </span>
                        </div>
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_COHESION_PLOT__
                    </div>

                    <div class="chart-card" style="grid-column: 1 / -1; height: auto; padding: 20px;">
                        __HTML_FACTION_LIST__
                    </div>
                </div>

                <script>
                    setTimeout(() => {
                        const myPlot = document.getElementById('network_graph');
                        if(!myPlot) return;

                        myPlot.on('plotly_hover', function(data){
                            const nodeTraceNum = myPlot.data.length - 1; 
                            const hoverCurve = data.points[0].curveNumber;

                            if(hoverCurve === nodeTraceNum) {
                                const pointIndex = data.points[0].pointIndex;
                                const neighbors = data.points[0].customdata; 
                                const nodeNames = myPlot.data[nodeTraceNum].text; 

                                let nodeOpacities = new Array(nodeNames.length).fill(0.1);
                                nodeOpacities[pointIndex] = 1.0; 

                                if(neighbors) {
                                    neighbors.forEach(neighbor => {
                                        const nIdx = nodeNames.indexOf(neighbor);
                                        if(nIdx > -1) nodeOpacities[nIdx] = 1.0;
                                    });
                                }

                                Plotly.restyle(myPlot, {'marker.opacity': [nodeOpacities]}, [nodeTraceNum]);
                            }
                        });

                        myPlot.on('plotly_unhover', function(data){
                            const nodeTraceNum = myPlot.data.length - 1;
                            const hoverCurve = data.points[0].curveNumber;
                            if(hoverCurve === nodeTraceNum) {
                                Plotly.restyle(myPlot, {'marker.opacity': 1.0}, [nodeTraceNum]);
                            }
                        });
                    }, 2000);
                </script>
            </body>
            </html>
            """

            dashboard_html = dashboard_html.replace("__MEETING_NAME__", str(meeting_name))
            dashboard_html = dashboard_html.replace("__TOTAL_TDOCS__", str(total_tdocs))
            dashboard_html = dashboard_html.replace("__TOTAL_COMPANIES__", str(total_companies))
            dashboard_html = dashboard_html.replace("__HTML_AI__", html_ai)
            dashboard_html = dashboard_html.replace("__HTML_STATUS__", html_status)
            dashboard_html = dashboard_html.replace("__HTML_COMP__", html_comp)
            dashboard_html = dashboard_html.replace("__HTML_NET__", html_net)
            dashboard_html = dashboard_html.replace("__HTML_CLUSTER_CONTRIBS__", html_cluster_contribs)
            dashboard_html = dashboard_html.replace("__HTML_COHESION_PLOT__", html_cohesion_plot)
            dashboard_html = dashboard_html.replace("__HTML_FACTION_LIST__", html_faction_list)

            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_html)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))