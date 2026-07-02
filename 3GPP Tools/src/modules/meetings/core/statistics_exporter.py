# --- File: src/modules/meetings/core/statistics_exporter.py ---
import os
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import networkx as nx

from modules.meetings.core.company_sanitizer import CompanySanitizer


class StatisticsExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, meeting_dir: Path, tdocs_data: list, mtg_info: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.tdocs_data = tdocs_data
        self.mtg_info = mtg_info
        self.export_dir = self.meeting_dir / "Export"

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            df = pd.DataFrame(self.tdocs_data)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            # --- Data Sanitization ---
            df = df[~df['TDoc Status'].str.lower().str.contains('withdrawn', na=False)].copy()
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            # --- 1. TDocs per AI (Bar Chart) ---
            ai_counts = df['Agenda Item'].value_counts().reset_index()
            ai_counts.columns = ['Agenda Item', 'Count']
            ai_counts = ai_counts[ai_counts['Agenda Item'].str.strip() != '']
            fig_ai = px.bar(ai_counts.head(20), x='Agenda Item', y='Count', title="Top 20 Agenda Items by TDoc Volume",
                            color_discrete_sequence=['#005A9E'])
            fig_ai.write_html(str(self.export_dir / "Stat_AI_Volume.html"))
            html_ai = fig_ai.to_html(full_html=False, include_plotlyjs='cdn', default_height="100%",
                                     default_width="100%")

            # --- 2. TDocs by Status (Donut Chart) ---
            status_counts = df['TDoc Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            status_counts = status_counts[status_counts['Status'].str.strip() != '']
            fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.4, title="TDoc Outcomes")
            fig_status.update_traces(textposition='inside', textinfo='percent+label')
            fig_status.write_html(str(self.export_dir / "Stat_Outcomes.html"))
            html_status = fig_status.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                             default_width="100%")

            # --- 3. Top Contributors (Horizontal Bar Chart) ---
            all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
            comp_counts = pd.Series(all_companies).value_counts().reset_index()
            comp_counts.columns = ['Company', 'Count']
            fig_comp = px.bar(comp_counts.head(30).sort_values('Count', ascending=True),
                              x='Count', y='Company', orientation='h',
                              title="Top 30 Contributing Companies", color_discrete_sequence=['#107C10'])
            fig_comp.write_html(str(self.export_dir / "Stat_Top_Contributors.html"))
            html_comp = fig_comp.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                         default_width="100%")

            # --- 4. Co-Signing Network Graph (NetworkX + Plotly) ---
            G = nx.Graph()
            for companies in df['Clean_Companies']:
                if len(companies) > 1:
                    for i in range(len(companies)):
                        for j in range(i + 1, len(companies)):
                            c1, c2 = companies[i], companies[j]
                            if G.has_edge(c1, c2):
                                G[c1][c2]['weight'] += 1
                            else:
                                G.add_edge(c1, c2, weight=1)

            # Keep threshold at 1 to capture ALL relations
            edges_to_remove = [(u, v) for u, v, data in G.edges(data=True) if data['weight'] < 1]
            G.remove_edges_from(edges_to_remove)
            G.remove_nodes_from(list(nx.isolates(G)))

            html_net = "<p style='padding:20px; color:#666;'>Not enough co-signed documents to generate network graph.</p>"
            html_cluster = ""

            if len(G.nodes) > 0:
                # Louvain Community Detection
                communities = list(nx.community.louvain_communities(G, seed=42))
                community_map = {}
                cluster_names = {}

                for i, comm in enumerate(communities):
                    dominant_node = max(comm, key=lambda n: G.degree(n))
                    cluster_names[i] = f"Cluster: {dominant_node} & Allies"
                    for node in comm:
                        community_map[node] = i

                pos = nx.spring_layout(G, k=0.5, seed=42)
                traces = []

                # Proportional Edge Thickness Grouped by Weight
                max_weight = max([data['weight'] for u, v, data in G.edges(data=True)]) if G.edges else 1
                edge_weights = set([data['weight'] for u, v, data in G.edges(data=True)])

                for weight in edge_weights:
                    edge_x, edge_y = [], []
                    for u, v, data in G.edges(data=True):
                        if data['weight'] == weight:
                            edge_x.extend([pos[u][0], pos[v][0], None])
                            edge_y.extend([pos[u][1], pos[v][1], None])

                    calc_width = 0.5 + (weight / max_weight) * 4.0
                    hover_label = f"Co-signed {weight} docs"

                    traces.append(go.Scatter(x=edge_x, y=edge_y, line=dict(width=calc_width, color='#999'),
                                             hoverinfo='name', name=hover_label, mode='lines', opacity=0.6))

                # Node Data Preparation
                node_x, node_y, node_text, node_size, node_color, custom_data = [], [], [], [], [], []
                for node in G.nodes():
                    x, y = pos[node]
                    node_x.append(x)
                    node_y.append(y)

                    adj_size = len(list(G.neighbors(node))) * 3
                    node_size.append(max(10, min(adj_size, 50)))
                    node_color.append(community_map[node])

                    neighbors = list(G.neighbors(node))
                    custom_data.append(neighbors)

                    hover_info = f"<b>{node}</b><br>Faction: {cluster_names[community_map[node]]}<br>Partners: {len(neighbors)}<br>---<br>"
                    neighbor_weights = [(n, G[node][n]['weight']) for n in neighbors]
                    neighbor_weights.sort(key=lambda item: item[1], reverse=True)

                    for neighbor, weight in neighbor_weights[:15]:
                        hover_info += f"• {neighbor} ({weight} shared)<br>"
                    if len(neighbors) > 15: hover_info += f"<i>...and {len(neighbors) - 15} more</i>"

                    node_text.append(hover_info)

                # Append the node trace LAST so it renders on top of the edges
                # FIXED: Changed colorscale to 'turbo'
                node_trace = go.Scatter(x=node_x, y=node_y, mode='markers+text', text=list(G.nodes()),
                                        textposition="top center", hovertext=node_text, hoverinfo='text',
                                        customdata=custom_data, name="Companies",
                                        marker=dict(showscale=False, colorscale='turbo', size=node_size,
                                                    color=node_color, line_width=1, line_color='#fff'))
                traces.append(node_trace)

                fig_net = go.Figure(data=traces,
                                    layout=go.Layout(title='Strategic Co-Signing Alliances (Network Graph)',
                                                     showlegend=False, hovermode='closest',
                                                     margin=dict(b=20, l=5, r=5, t=40),
                                                     xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                                                     yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)))

                fig_net.write_html(str(self.export_dir / "Stat_Network_Alliances.html"))
                html_net = fig_net.to_html(full_html=False, include_plotlyjs=False, div_id="network_graph",
                                           default_height="100%", default_width="100%")

                # --- 5. Community / Factions Chart ---
                comm_sizes = pd.Series(
                    [cluster_names[community_map[n]] for n in G.nodes()]).value_counts().reset_index()
                comm_sizes.columns = ['Faction', 'Members']
                fig_cluster = px.bar(comm_sizes, x='Members', y='Faction', orientation='h',
                                     title="Detected Co-Signing Factions (Louvain Algorithm)", color='Faction')
                fig_cluster.write_html(str(self.export_dir / "Stat_Factions.html"))
                html_cluster = fig_cluster.to_html(full_html=False, include_plotlyjs=False, default_height="100%",
                                                   default_width="100%")

            # --- 6. Compile Dashboard HTML (Clean Standard String, No CSS/JS Escaping Required) ---
            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"
            total_tdocs = len(df)
            total_companies = len(comp_counts)

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
                    .kpi-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 200px; border-top: 4px solid #0078D7; }
                    .kpi-card h3 { margin: 0; font-size: 32px; color: #0078D7; }
                    .kpi-card p { margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }
                    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }

                    /* Dynamic Flexbox Card Styling */
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
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_AI__
                    </div>
                    <div class="chart-card">
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_STATUS__
                    </div>
                    <div class="chart-card" style="grid-column: 1 / -1; height: 600px;">
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_COMP__
                    </div>
                    <div class="chart-card" style="grid-column: 1 / -1; height: 750px;">
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_NET__
                    </div>
                    <div class="chart-card" style="grid-column: 1 / -1; height: 400px;">
                        <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                        __HTML_CLUSTER__
                    </div>
                </div>

                <!-- JS Engine to intercept Plotly Hovers and fade unassociated nodes -->
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

            # Safely inject the data using string replacement
            dashboard_html = dashboard_html.replace("__MEETING_NAME__", str(meeting_name))
            dashboard_html = dashboard_html.replace("__TOTAL_TDOCS__", str(total_tdocs))
            dashboard_html = dashboard_html.replace("__TOTAL_COMPANIES__", str(total_companies))
            dashboard_html = dashboard_html.replace("__HTML_AI__", html_ai)
            dashboard_html = dashboard_html.replace("__HTML_STATUS__", html_status)
            dashboard_html = dashboard_html.replace("__HTML_COMP__", html_comp)
            dashboard_html = dashboard_html.replace("__HTML_NET__", html_net)
            dashboard_html = dashboard_html.replace("__HTML_CLUSTER__", html_cluster)

            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_html)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))