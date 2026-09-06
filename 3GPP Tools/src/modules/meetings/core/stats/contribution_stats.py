import os
import re
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import networkx as nx

from core.config.plot_styles import PALETTE, THEME_COLOR, CLUSTER_PALETTE
from core.utils.company_sanitizer import CompanySanitizer


class ContributionStatsExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        export_dir: Path,
        tdocs_data: list,
        target_companies: set,
        target_wis: list,
        config: dict,
        parent=None
    ):
        super().__init__(parent)
        self.export_dir = Path(export_dir)
        self.tdocs_data = tdocs_data
        self.target_companies = set(target_companies) if target_companies else set()
        self.target_wis = target_wis or []
        self.config = config or {}

        self.threshold = self.config.get("threshold", 1)
        self.resolution = self.config.get("resolution", 1.5)

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            df = pd.DataFrame(self.tdocs_data)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            # Exclude withdrawn documents from statistical analysis
            df = df[~df['TDoc Status'].str.lower().str.contains('withdrawn', na=False)].copy()
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            # --- 1. Extract Month (YYYY-MM) from Meeting End Date ---
            def extract_month(date_str):
                if not date_str or not isinstance(date_str, str):
                    return "Undated"
                match = re.search(r'(\d{4})[-/.](\d{2})', date_str)
                return f"{match.group(1)}-{match.group(2)}" if match else "Undated"

            df['Month'] = df.get('end_date', pd.Series(dtype='object')).apply(extract_month)

            # --- 2. Extract Work Item ---
            def extract_wi(row):
                for k in ['Work Item', 'WI', 'WID', 'Work Item / Study Item']:
                    val = str(row.get(k, '')).strip()
                    if val and val.lower() not in ['', 'none', 'unknown', '-']:
                        return val
                return "Unspecified"

            df['Clean_WI'] = df.apply(extract_wi, axis=1)

            # --- KPI Metrics Calculation ---
            total_tdocs = len(df)
            agreed_count = df['TDoc Status'].str.lower().str.contains('agreed|approved', na=False).sum()
            agree_pct = round((agreed_count / total_tdocs) * 100, 1) if total_tdocs else 0

            solo_count = sum(1 for comps in df['Clean_Companies'] if len(comps) <= 1)
            joint_count = total_tdocs - solo_count
            joint_pct = round((joint_count / total_tdocs) * 100, 1) if total_tdocs else 0

            # --- PLOT 1: Monthly Timeline Trend ---
            html_timeline = self._generate_timeline_plot(df)

            # --- PLOT 2: Overall Outcomes ---
            html_status = self._generate_outcomes_plot(df)

            # --- PLOT 3: Work Item Allocation ---
            html_wi = self._generate_wi_allocation_plot(df)

            # --- PLOT 4: Top Third-Party Partner Vendors ---
            html_partners = self._generate_partner_vendors_plot(df)

            # --- PLOT 5: Alliance Network (Third-Party Highlighted) ---
            html_network = self._generate_alliance_network_plot(df)

            # --- PLOT 6: Company vs Topic Focus Matrix (Heatmap) ---
            html_heatmap = self._generate_heatmap_plot(df)

            # --- Assemble Dashboard ---
            target_str = ", ".join(sorted(list(self.target_companies))) if self.target_companies else "All Entities"
            dashboard_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>3GPP Contribution & Strategic Alliances Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #F8FAFC; margin: 0; padding: 24px; color: #1E293B; }}
        h1 {{ color: #0F172A; text-align: center; margin: 0 0 6px 0; font-size: 26px; }}
        p.subtitle {{ text-align: center; color: #64748B; margin: 0 0 24px 0; font-size: 13px; }}
        .kpi-container {{ display: flex; justify-content: center; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }}
        .kpi-card {{ background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.06); padding: 16px 20px; text-align: center; min-width: 170px; border-top: 4px solid #1E5C99; }}
        .kpi-card h3 {{ margin: 0; font-size: 28px; color: #1E5C99; }}
        .kpi-card p {{ margin: 4px 0 0; color: #64748B; font-size: 11px; text-transform: uppercase; font-weight: bold; }}
        .grid-container {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .chart-card {{ position: relative; background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.06); padding: 40px 16px 16px 16px; height: 480px; display: flex; flex-direction: column; }}
        .chart-card > div {{ flex-grow: 1; width: 100%; height: 100%; }}
        .fs-btn {{ position: absolute; top: 12px; right: 12px; z-index: 100; cursor: pointer; background: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE; border-radius: 4px; padding: 4px 10px; font-weight: bold; font-size: 11px; }}
        .fs-btn:hover {{ background: #DBEAFE; }}
        .chart-card.fullscreen {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 9999; margin: 0; border-radius: 0; padding: 50px 20px 20px 20px; box-sizing: border-box; }}
    </style>
    <script>
        function toggleFullscreen(btn) {{
            const card = btn.parentElement;
            card.classList.toggle('fullscreen');
            btn.innerHTML = card.classList.contains('fullscreen') ? '✖ Close' : '⛶ Expand';
            setTimeout(() => {{ window.dispatchEvent(new Event('resize')); }}, 60);
        }}
    </script>
</head>
<body>
    <h1>📊 3GPP Cross-Meeting Contribution & Alliance Audit</h1>
    <p class="subtitle">Focus Entities: <b>{target_str}</b> | Audited Dataset: <b>{total_tdocs} Contributions</b></p>

    <div class="kpi-container">
        <div class="kpi-card"><h3>{total_tdocs}</h3><p>Total TDocs</p></div>
        <div class="kpi-card"><h3>{agree_pct}%</h3><p>Agreement Rate</p></div>
        <div class="kpi-card"><h3>{joint_pct}%</h3><p>Joint Contributions</p></div>
        <div class="kpi-card"><h3>{solo_count}</h3><p>Solo Submissions</p></div>
        <div class="kpi-card"><h3>{df['Clean_WI'].nunique()}</h3><p>Active Work Items</p></div>
    </div>

    <div class="grid-container">
        <div class="chart-card" style="grid-column: 1 / -1; height: 460px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_timeline}
        </div>

        <div class="chart-card">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_status}
        </div>

        <div class="chart-card">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_partners}
        </div>

        <div class="chart-card" style="grid-column: 1 / -1; height: 500px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_wi}
        </div>

        <div class="chart-card" style="grid-column: 1 / -1; height: 720px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_network}
        </div>

        <div class="chart-card" style="grid-column: 1 / -1; height: 600px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_heatmap}
        </div>
    </div>
</body>
</html>
"""

            out_file = self.export_dir / "Contributions_Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_html)

            self.finished.emit(True, str(out_file))

        except Exception as e:
            self.finished.emit(False, str(e))

    def _generate_timeline_plot(self, df: pd.DataFrame) -> str:
        timeline_df = df[df['Month'] != 'Undated'].copy()
        if timeline_df.empty:
            timeline_df = df.copy()
            timeline_df['Month'] = timeline_df.get('Meeting', 'All')

        grouped = timeline_df.groupby(['Month', 'TDoc Status']).size().reset_index(name='Count')
        grouped = grouped.sort_values('Month')

        fig = px.bar(
            grouped,
            x='Month',
            y='Count',
            color='TDoc Status',
            title="Monthly Contribution Trend by Outcome Status (End Date / Month)",
            color_discrete_sequence=PALETTE,
            barmode='stack'
        )
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend_title_text="")
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_outcomes_plot(self, df: pd.DataFrame) -> str:
        status_counts = df['TDoc Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        status_counts = status_counts[status_counts['Status'].str.strip() != '']

        fig = px.pie(
            status_counts,
            names='Status',
            values='Count',
            hole=0.4,
            title="Overall Contribution Outcomes",
            color_discrete_sequence=PALETTE
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_wi_allocation_plot(self, df: pd.DataFrame) -> str:
        wi_counts = df['Clean_WI'].value_counts().reset_index()
        wi_counts.columns = ['Work Item', 'Count']
        plot_df = wi_counts.head(20).sort_values('Count', ascending=True)

        fig = px.bar(
            plot_df,
            x='Count',
            y='Work Item',
            orientation='h',
            title="Work Item & Study Item Allocation (Top 20)",
            color_discrete_sequence=[THEME_COLOR]
        )
        fig.update_yaxes(tickmode='linear', dtick=1, title=None)
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_partner_vendors_plot(self, df: pd.DataFrame) -> str:
        partner_counts = {}
        for companies in df['Clean_Companies']:
            if self.target_companies:
                targets_in_doc = [c for c in companies if c in self.target_companies]
                if targets_in_doc:
                    for c in companies:
                        if c not in self.target_companies:
                            partner_counts[c] = partner_counts.get(c, 0) + 1
            else:
                for c in companies:
                    partner_counts[c] = partner_counts.get(c, 0) + 1

        if not partner_counts:
            return "<p style='padding:20px; color:#666;'>No third-party co-authors discovered in the filtered dataset.</p>"

        plot_df = pd.DataFrame(list(partner_counts.items()), columns=['Vendor', 'Joint TDocs'])
        plot_df = plot_df.sort_values('Joint TDocs', ascending=True).tail(15)

        title = "Top Co-Signing Third-Party Vendors" if self.target_companies else "Top Contributing Entities"
        fig = px.bar(
            plot_df,
            x='Joint TDocs',
            y='Vendor',
            orientation='h',
            title=title,
            color_discrete_sequence=["#0D9488"]
        )
        fig.update_yaxes(tickmode='linear', dtick=1, title=None)
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_alliance_network_plot(self, df: pd.DataFrame) -> str:
        G = nx.Graph()
        for companies in df['Clean_Companies']:
            if len(companies) > 1:
                for i in range(len(companies)):
                    for j in range(i + 1, len(companies)):
                        c1, c2 = companies[i], companies[j]
                        if self.target_companies and (c1 not in self.target_companies and c2 not in self.target_companies):
                            continue
                        if G.has_edge(c1, c2):
                            G[c1][c2]['weight'] += 1
                        else:
                            G.add_edge(c1, c2, weight=1)

        edges_to_remove = [(u, v) for u, v, data in G.edges(data=True) if data['weight'] < self.threshold]
        G.remove_edges_from(edges_to_remove)
        G.remove_nodes_from(list(nx.isolates(G)))

        if len(G.nodes) == 0:
            return "<p style='padding:20px; color:#666;'>Not enough co-signed documents to construct the alliance network graph.</p>"

        pos = nx.spring_layout(G, k=0.6, seed=42)
        traces = []

        # Draw Edges
        max_weight = max([data['weight'] for u, v, data in G.edges(data=True)]) if G.edges else 1
        edge_x, edge_y = [], []
        mid_x, mid_y, mid_text = [], [], []

        for u, v, data in G.edges(data=True):
            edge_x.extend([pos[u][0], pos[v][0], None])
            edge_y.extend([pos[u][1], pos[v][1], None])
            mid_x.append((pos[u][0] + pos[v][0]) / 2)
            mid_y.append((pos[u][1] + pos[v][1]) / 2)
            mid_text.append(f"<b>{u}</b> 🤝 <b>{v}</b><br>Shared TDocs: {data['weight']}")

        traces.append(go.Scatter(
            x=edge_x, y=edge_y, line=dict(width=1.5, color='#CBD5E1'),
            hoverinfo='none', mode='lines', opacity=0.7
        ))

        traces.append(go.Scatter(
            x=mid_x, y=mid_y, mode='markers', hovertext=mid_text,
            hovertemplate="%{hovertext}<extra></extra>",
            marker=dict(size=12, color='rgba(255,255,255,0.01)', line=dict(width=0)),
            showlegend=False
        ))

        # Draw Nodes with Target vs Third-Party Highlighting
        node_x, node_y, node_text, node_size, node_color, node_symbols = [], [], [], [], [], []

        for node in G.nodes():
            node_x.append(pos[node][0])
            node_y.append(pos[node][1])
            is_target = node in self.target_companies
            neighbors = list(G.neighbors(node))

            if is_target:
                node_size.append(26)
                node_color.append("#E20074")  # Magenta highlight for target
                node_symbols.append("diamond")
                role_label = "Target Entity"
            else:
                node_size.append(max(12, min(len(neighbors) * 3, 22)))
                node_color.append("#0284C7")  # Partner blue
                node_symbols.append("circle")
                role_label = "Third-Party Vendor"

            hover_info = f"<b>[{role_label}] {node}</b><br>Connected Allies: {len(neighbors)}<br><br><b>Top Collaborations:</b><br>"
            neighbor_weights = sorted([(n, G[node][n]['weight']) for n in neighbors], key=lambda x: x[1], reverse=True)
            for neighbor, weight in neighbor_weights[:8]:
                hover_info += f"• {neighbor} ({weight} shared)<br>"
            node_text.append(hover_info)

        traces.append(go.Scatter(
            x=node_x, y=node_y, mode='markers+text', text=list(G.nodes()),
            textposition="top center", hovertext=node_text,
            hovertemplate="%{hovertext}<extra></extra>", name="Entities",
            marker=dict(size=node_size, color=node_color, symbol=node_symbols, line_width=1.5, line_color='#FFFFFF')
        ))

        fig = go.Figure(data=traces, layout=go.Layout(
            title="Strategic Alliance Network (Highlighted: Target Companies ⬥ vs. Third-Party Vendors ●)",
            showlegend=False, hovermode='closest', margin=dict(b=20, l=5, r=5, t=40),
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
        ))
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_heatmap_plot(self, df: pd.DataFrame) -> str:
        exploded = df.explode('Clean_Companies')
        exploded = exploded.dropna(subset=['Clean_Companies', 'Agenda Item'])
        exploded = exploded[(exploded['Clean_Companies'].str.strip() != '') & (exploded['Agenda Item'].str.strip() != '')]

        if exploded.empty:
            return "<p style='padding:20px; color:#666;'>No data available for company heatmap.</p>"

        top_comps = exploded['Clean_Companies'].value_counts().head(20).index
        top_ais = exploded['Agenda Item'].value_counts().head(20).index

        filtered = exploded[exploded['Clean_Companies'].isin(top_comps) & exploded['Agenda Item'].isin(top_ais)]
        matrix = pd.crosstab(filtered['Clean_Companies'], filtered['Agenda Item'])

        if matrix.empty:
            return "<p style='padding:20px; color:#666;'>No data available for company heatmap matrix.</p>"

        matrix = matrix.loc[matrix.sum(axis=1).sort_values(ascending=False).index]
        matrix = matrix[matrix.sum(axis=0).sort_values(ascending=False).index]

        fig = px.imshow(
            matrix,
            labels=dict(x="Agenda Item", y="Company", color="TDocs"),
            x=matrix.columns,
            y=matrix.index,
            text_auto=True,
            aspect="auto",
            title="Company Focus Matrix (Top Companies vs. Top Topics)",
            color_continuous_scale="Blues"
        )
        fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")