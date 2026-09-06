# --- File: src/modules/meetings/core/stats/contribution_stats.py ---
import os
import re
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
import plotly.graph_objects as go

from core.config.plot_styles import THEME_COLOR
from core.utils.company_sanitizer import CompanySanitizer

# Semantic color palette for standard 3GPP contribution outcomes
STATUS_COLORS = {
    'Agreed': '#10B981',       # Emerald Green
    'Approved': '#059669',     # Deep Green
    'Not Treated': '#94A3B8',  # Slate Grey
    'Revised': '#6366F1',      # Indigo Blue
    'Noted': '#F59E0B',        # Amber
    'Postponed': '#EAB308',    # Yellow
    'Merged': '#06B6D4',       # Cyan
    'Endorsed': '#14B8A6',     # Teal
    'Replied To': '#8B5CF6',   # Purple
    'Available': '#38BDF8',    # Sky Blue
    'Withdrawn': '#EF4444',    # Red
    'Unknown': '#CBD5E1'       # Muted Grey
}

STATUS_ORDER = [
    'Agreed', 'Approved', 'Merged', 'Endorsed',
    'Noted', 'Postponed', 'Replied To', 'Available',
    'Not Treated', 'Revised', 'Unknown'
]


def normalize_status(val: str) -> str:
    """Normalizes raw 3GPP outcome strings into canonical categories."""
    s = str(val).strip().lower()
    if not s or s in ['nan', 'none', '-', '']:
        return 'Unknown'
    if 'not treated' in s or 'not handled' in s:
        return 'Not Treated'
    if 'agreed' in s:
        return 'Agreed'
    if 'approved' in s:
        return 'Approved'
    if 'revised' in s:
        return 'Revised'
    if 'noted' in s:
        return 'Noted'
    if 'postponed' in s:
        return 'Postponed'
    if 'merged' in s:
        return 'Merged'
    if 'endorsed' in s:
        return 'Endorsed'
    if 'replied' in s:
        return 'Replied To'
    if 'available' in s:
        return 'Available'
    if 'withdrawn' in s:
        return 'Withdrawn'
    return s.title()


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

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            df = pd.DataFrame(self.tdocs_data)
            if df.empty:
                self.finished.emit(False, "No TDoc data available to generate statistics.")
                return

            # Canonical status mapping
            df['Clean_Status'] = df['TDoc Status'].apply(normalize_status)

            # Exclude withdrawn documents from statistical analysis
            df = df[df['Clean_Status'] != 'Withdrawn'].copy()
            df['Clean_Companies'] = df['Source'].apply(CompanySanitizer.get_matching_contributors)

            # Working Group cleanup
            df['WG'] = df['WG'].fillna('').replace('', '3GPP')

            # Extract Month (YYYY-MM) from Meeting End Date
            def extract_month(date_str):
                if not date_str or not isinstance(date_str, str):
                    return "Undated"
                match = re.search(r'(\d{4})[-/.](\d{2})', date_str)
                return f"{match.group(1)}-{match.group(2)}" if match else "Undated"

            df['Month'] = df.get('end_date', pd.Series(dtype='object')).apply(extract_month)

            # Extract list of Related WIs (handles comma-separated multi-entries)
            def extract_wi_list(row):
                for k in ['Related WIs', 'Work Item', 'WI', 'WID']:
                    val = str(row.get(k, '')).strip()
                    if val and val.lower() not in ['', 'none', 'unknown', '-', 'dummy']:
                        items = [w.strip() for w in re.split(r'[,;]+', val) if w.strip() and w.strip().lower() != 'dummy']
                        if items:
                            return items
                return ["Unspecified"]

            df['Clean_WI_List'] = df.apply(extract_wi_list, axis=1)

            total_tdocs = len(df)
            agreed_approved_count = df['Clean_Status'].isin(['Agreed', 'Approved']).sum()
            gross_agree_pct = round((agreed_approved_count / total_tdocs) * 100, 1) if total_tdocs else 0

            # Consensus Rate: Success among completed decisions (excluding not-treated and intermediate revisions)
            decided_success = df['Clean_Status'].isin(['Agreed', 'Approved', 'Merged', 'Endorsed']).sum()
            total_decided = df['Clean_Status'].isin(['Agreed', 'Approved', 'Merged', 'Endorsed', 'Noted', 'Postponed']).sum()
            consensus_pct = round((decided_success / total_decided) * 100, 1) if total_decided else 0

            solo_count = sum(1 for comps in df['Clean_Companies'] if len(comps) <= 1)
            joint_count = total_tdocs - solo_count
            joint_pct = round((joint_count / total_tdocs) * 100, 1) if total_tdocs else 0

            unique_wgs = df['WG'].nunique()
            unique_wis = len(set(wi for sublist in df['Clean_WI_List'] for wi in sublist if wi not in ['Unspecified', 'DUMMY']))

            # Generate Charts using native Python int arrays (eliminates binary bdata)
            html_timeline = self._generate_timeline_plot(df)
            html_status = self._generate_outcomes_plot(df)
            html_partners = self._generate_partner_vendors_plot(df)
            html_wg_activity = self._generate_wg_activity_plot(df)
            html_wi = self._generate_wi_allocation_plot(df)

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
        .chart-card {{ position: relative; background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.06); padding: 40px 16px 16px 16px; height: 460px; display: flex; flex-direction: column; }}
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
    <h1>📊 3GPP Cross-Meeting Contribution & Working Group Audit</h1>
    <p class="subtitle">Focus Entities: <b>{target_str}</b> | Audited Dataset: <b>{total_tdocs} Contributions</b></p>

    <div class="kpi-container">
        <div class="kpi-card"><h3>{total_tdocs}</h3><p>Total TDocs</p></div>
        <div class="kpi-card"><h3>{consensus_pct}%</h3><p>Consensus Rate ({decided_success}/{total_decided} Decided)</p></div>
        <div class="kpi-card"><h3>{gross_agree_pct}%</h3><p>Gross Agreement ({agreed_approved_count}/{total_tdocs} Gross)</p></div>
        <div class="kpi-card"><h3>{unique_wgs}</h3><p>Active Working Groups</p></div>
        <div class="kpi-card"><h3>{joint_pct}%</h3><p>Joint Contributions</p></div>
        <div class="kpi-card"><h3>{unique_wis}</h3><p>Active Work Items</p></div>
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

        <div class="chart-card" style="grid-column: 1 / -1; height: 480px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_wg_activity}
        </div>

        <div class="chart-card" style="grid-column: 1 / -1; height: 480px;">
            <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
            {html_wi}
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

        grouped = timeline_df.groupby(['Month', 'Clean_Status']).size().reset_index(name='Count')
        pivot = grouped.pivot(index='Month', columns='Clean_Status', values='Count').fillna(0)
        months = sorted(pivot.index.tolist())

        fig = go.Figure()
        for status in STATUS_ORDER:
            if status in pivot.columns:
                y_vals = [int(pivot.loc[m, status]) for m in months]
                if sum(y_vals) > 0:
                    fig.add_trace(go.Bar(
                        name=status,
                        x=months,
                        y=y_vals,
                        marker=dict(color=STATUS_COLORS.get(status, '#94A3B8')),
                        hovertemplate=f"<b>{status}</b><br>Month: %{{x}}<br>TDocs: %{{y}}<extra></extra>"
                    ))

        fig.update_layout(
            barmode='stack',
            title="Monthly Contribution Trend by Outcome Status (End Date / Month)",
            xaxis_title="Meeting Month",
            yaxis_title="TDocs Count",
            legend_title_text="Outcome",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_outcomes_plot(self, df: pd.DataFrame) -> str:
        counts = df['Clean_Status'].value_counts()
        counts = counts[counts.index.str.strip() != '']

        labels = [s for s in STATUS_ORDER if s in counts.index]
        values = [int(counts[s]) for s in labels]
        colors = [STATUS_COLORS.get(s, '#94A3B8') for s in labels]

        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.42,
            sort=False,
            marker=dict(colors=colors, line=dict(color='#FFFFFF', width=1.5)),
            textinfo='percent+label',
            hovertemplate="<b>%{label}</b><br>TDocs: %{value}<br>Percentage: %{percent}<extra></extra>"
        )])

        fig.update_layout(
            title="Overall Contribution Outcomes",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02)
        )
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

        sorted_partners = sorted(partner_counts.items(), key=lambda x: x[1])[-15:]
        vendors = [p[0] for p in sorted_partners]
        counts = [int(p[1]) for p in sorted_partners]

        title = "Top Co-Signing Third-Party Vendors" if self.target_companies else "Top Contributing Entities"
        fig = go.Figure(go.Bar(
            x=counts,
            y=vendors,
            orientation='h',
            marker=dict(color="#0D9488"),
            hovertemplate="<b>%{y}</b><br>Joint TDocs: %{x}<extra></extra>"
        ))
        fig.update_layout(
            title=title,
            xaxis_title="Joint Contributions",
            yaxis_title=None,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_wg_activity_plot(self, df: pd.DataFrame) -> str:
        wg_df = df[df['WG'].str.strip() != ''].copy()
        if wg_df.empty:
            return "<p style='padding:20px; color:#666;'>No Working Group data available.</p>"

        grouped = wg_df.groupby(['WG', 'Clean_Status']).size().reset_index(name='Count')
        pivot = grouped.pivot(index='WG', columns='Clean_Status', values='Count').fillna(0)
        wgs = sorted(pivot.index.tolist(), key=lambda w: wg_df['WG'].value_counts().get(w, 0), reverse=True)

        fig = go.Figure()
        for status in STATUS_ORDER:
            if status in pivot.columns:
                y_vals = [int(pivot.loc[w, status]) for w in wgs]
                if sum(y_vals) > 0:
                    fig.add_trace(go.Bar(
                        name=status,
                        x=wgs,
                        y=y_vals,
                        marker=dict(color=STATUS_COLORS.get(status, '#94A3B8')),
                        hovertemplate=f"<b>{status}</b><br>WG: %{{x}}<br>TDocs: %{{y}}<extra></extra>"
                    ))

        fig.update_layout(
            barmode='stack',
            title="Working Group Activity by Outcome Status",
            xaxis_title="Working Group",
            yaxis_title="Contributions Count",
            legend_title_text="Outcome",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    def _generate_wi_allocation_plot(self, df: pd.DataFrame) -> str:
        exploded_wi = df.explode('Clean_WI_List')
        wi_counts = exploded_wi['Clean_WI_List'].value_counts()
        # Filter out Unspecified and DUMMY placeholders
        wi_counts = wi_counts[~wi_counts.index.isin(['Unspecified', 'DUMMY'])].head(20).sort_values(ascending=True)

        if wi_counts.empty:
            wi_counts = exploded_wi['Clean_WI_List'].value_counts().head(20).sort_values(ascending=True)

        wis = wi_counts.index.tolist()
        counts = [int(v) for v in wi_counts.values.tolist()]

        fig = go.Figure(go.Bar(
            x=counts,
            y=wis,
            orientation='h',
            marker=dict(color=THEME_COLOR),
            hovertemplate="<b>%{y}</b><br>TDocs: %{x}<extra></extra>"
        ))
        fig.update_layout(
            title="Work Item & Study Item Allocation (Top 20 Related WIs)",
            xaxis_title="Contributions Count",
            yaxis_title=None,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        return fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")