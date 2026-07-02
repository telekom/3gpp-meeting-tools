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
            # Remove entirely withdrawn documents
            df = df[~df['TDoc Status'].str.lower().str.contains('withdrawn', na=False)].copy()

            # Extract Canonical Companies
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            # --- 1. TDocs per AI (Bar Chart) ---
            ai_counts = df['Agenda Item'].value_counts().reset_index()
            ai_counts.columns = ['Agenda Item', 'Count']
            ai_counts = ai_counts[ai_counts['Agenda Item'].str.strip() != '']
            fig_ai = px.bar(ai_counts.head(20), x='Agenda Item', y='Count', title="Top 20 Agenda Items by TDoc Volume",
                            color_discrete_sequence=['#005A9E'])
            html_ai = fig_ai.to_html(full_html=False, include_plotlyjs='cdn')

            # --- 2. TDocs by Status (Donut Chart) ---
            status_counts = df['TDoc Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            status_counts = status_counts[status_counts['Status'].str.strip() != '']
            fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.4, title="TDoc Outcomes")
            fig_status.update_traces(textposition='inside', textinfo='percent+label')
            html_status = fig_status.to_html(full_html=False, include_plotlyjs=False)

            # --- 3. Top Contributors (Horizontal Bar Chart) ---
            all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
            comp_counts = pd.Series(all_companies).value_counts().reset_index()
            comp_counts.columns = ['Company', 'Count']
            fig_comp = px.bar(comp_counts.head(20).sort_values('Count', ascending=True),
                              x='Count', y='Company', orientation='h',
                              title="Top 20 Contributing Companies", color_discrete_sequence=['#107C10'])
            html_comp = fig_comp.to_html(full_html=False, include_plotlyjs=False)

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

            # Threshold filter: Only show connections if they co-signed at least 2 documents
            threshold = 2
            edges_to_remove = [(u, v) for u, v, data in G.edges(data=True) if data['weight'] < threshold]
            G.remove_edges_from(edges_to_remove)
            G.remove_nodes_from(list(nx.isolates(G)))

            if len(G.nodes) > 0:
                pos = nx.spring_layout(G, k=0.5, seed=42)

                edge_x, edge_y = [], []
                for edge in G.edges():
                    x0, y0 = pos[edge[0]]
                    x1, y1 = pos[edge[1]]
                    edge_x.extend([x0, x1, None])
                    edge_y.extend([y0, y1, None])

                edge_trace = go.Scatter(x=edge_x, y=edge_y, line=dict(width=0.5, color='#888'), hoverinfo='none',
                                        mode='lines')

                node_x, node_y, node_text, node_size = [], [], [], []
                for node in G.nodes():
                    x, y = pos[node]
                    node_x.append(x)
                    node_y.append(y)

                    # Node size scales with the total number of unique partners
                    adj_size = len(list(G.neighbors(node))) * 3
                    node_size.append(max(10, min(adj_size, 50)))

                    # --- THE FIX: Detailed HTML Hover Tooltip with Connection Weights ---
                    neighbors = list(G.neighbors(node))
                    hover_info = f"<b>{node}</b><br>Total Co-signing Partners: {len(neighbors)}<br>---<br>"

                    # Sort neighbors by weight descending so strongest alliances appear at the top
                    neighbor_weights = [(n, G[node][n]['weight']) for n in neighbors]
                    neighbor_weights.sort(key=lambda item: item[1], reverse=True)

                    for neighbor, weight in neighbor_weights:
                        hover_info += f"• {neighbor} ({weight} shared TDocs)<br>"

                    node_text.append(hover_info)

                node_trace = go.Scatter(x=node_x, y=node_y, mode='markers',
                                        hovertext=node_text, hoverinfo='text',
                                        marker=dict(showscale=True, colorscale='YlGnBu', size=node_size,
                                                    color=node_size, line_width=2))

                fig_net = go.Figure(data=[edge_trace, node_trace],
                                    layout=go.Layout(title='Co-Signing Alliances Network Graph (Threshold >= 2)',
                                                     showlegend=False, hovermode='closest',
                                                     margin=dict(b=20, l=5, r=5, t=40),
                                                     xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                                                     yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)))
                html_net = fig_net.to_html(full_html=False, include_plotlyjs=False)
            else:
                html_net = "<p style='padding:20px; color:#666;'>Not enough co-signed documents to generate network graph.</p>"

            # --- 5. Compile Dashboard HTML ---
            meeting_name = f"{self.mtg_info.get('wg_name', 'WG')} {self.mtg_info.get('meeting_number', '')}"
            total_tdocs = len(df)
            total_companies = len(comp_counts)

            dashboard_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>3GPP Statistics - {meeting_name}</title>
                <style>
                    body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }}
                    h1 {{ color: #333; text-align: center; margin-bottom: 30px; }}
                    .kpi-container {{ display: flex; justify-content: center; gap: 20px; margin-bottom: 40px; }}
                    .kpi-card {{ background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 200px; border-top: 4px solid #0078D7; }}
                    .kpi-card h3 {{ margin: 0; font-size: 32px; color: #0078D7; }}
                    .kpi-card p {{ margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }}
                    .grid-container {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
                    .chart-card {{ background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 15px; }}
                    .full-width {{ grid-column: 1 / -1; }}
                </style>
            </head>
            <body>
                <h1>📊 {meeting_name} - TDoc Statistics Dashboard</h1>

                <div class="kpi-container">
                    <div class="kpi-card"><h3>{total_tdocs}</h3><p>Total TDocs</p></div>
                    <div class="kpi-card"><h3>{total_companies}</h3><p>Participating Companies</p></div>
                </div>

                <div class="grid-container">
                    <div class="chart-card">{html_ai}</div>
                    <div class="chart-card">{html_status}</div>
                    <div class="chart-card">{html_comp}</div>
                    <div class="chart-card">{html_net}</div>
                </div>
            </body>
            </html>
            """

            out_file = self.export_dir / "Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_html)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))