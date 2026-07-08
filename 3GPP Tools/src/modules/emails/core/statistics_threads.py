import re
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from plotly import express as px


class EmailStatsExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, meeting_dir: Path, email_data: list, meeting_name: str = "Meeting"):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.email_data = email_data
        self.meeting_name = meeting_name
        self.export_dir = self.meeting_dir / "Export"
        self.THEME_COLOR = '#0078D7'

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)
            df = pd.DataFrame(self.email_data)

            if df.empty:
                self.finished.emit(False, "No email data available to generate statistics.")
                return

            df['date_received'] = pd.to_datetime(df['date_received'], utc=True, errors='coerce').dt.tz_localize(None)

            # ---> THE FIX: Force Title and Upper casing to completely eliminate string fragmentation
            df['agenda_item'] = df['agenda_item'].astype(str).str.strip().str.upper()
            df['company'] = df['company'].astype(str).str.strip().str.title()

            # Generate Global View
            g_html_ai = self._generate_ai_volume(df, "Global")
            g_html_comp = self._generate_company_volume(df, "Global")
            g_html_time = self._generate_timeline(df, "Global")
            g_html_table = self._generate_delegate_table(df, "Global")

            # Find Unique Agenda Items (sorted naturally)
            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

            raw_ais = df['agenda_item'].dropna().unique()
            unique_ais = sorted(
                [str(ai).strip() for ai in raw_ais if str(ai).strip() and str(ai).strip() != "Unknown AI"],
                key=natural_sort_key)

            views_html_buffer = []
            dropdown_options = ['<option value="global">🌐 Overall Email Analytics</option>']

            # 1. Compile Global Block
            views_html_buffer.append(self._compile_view_block(
                "global", len(df), df['sender_email'].nunique(),
                g_html_ai, g_html_comp, g_html_time, g_html_table, is_visible=True
            ))

            # 2. Compile Per-AI Blocks
            for idx, ai_name in enumerate(unique_ais):
                ai_df = df[df['agenda_item'].str.strip() == ai_name].copy()
                if ai_df.empty: continue

                safe_id = f"ai_{idx}"
                dropdown_options.append(
                    f'<option value="{safe_id}">📌 Agenda Item {ai_name} ({len(ai_df)} Emails)</option>')

                ai_html_comp = self._generate_company_volume(ai_df, safe_id)
                ai_html_time = self._generate_timeline(ai_df, safe_id)
                ai_html_table = self._generate_delegate_table(ai_df, safe_id)

                views_html_buffer.append(self._compile_view_block(
                    safe_id, len(ai_df), ai_df['sender_email'].nunique(),
                    None, ai_html_comp, ai_html_time, ai_html_table, is_visible=False
                ))

            # 3. Assemble Final Dashboard
            dashboard_template = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>📧 Email Analytics - {self.meeting_name}</title>
                <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
                <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
                <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
                <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
                <style>
                    body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }}
                    h1 {{ color: #333; text-align: center; margin-bottom: 10px; }}
                    .selector-container {{ display: flex; justify-content: center; margin-bottom: 30px; background: #FFF; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #E0E0E0; }}
                    .selector-container label {{ font-weight: bold; margin-right: 12px; align-self: center; color: #444; }}
                    select {{ padding: 8px 16px; border-radius: 6px; border: 1px solid #CCCCCC; font-size: 14px; font-weight: bold; color: #005A9E; outline: none; background: #F4F8FC; cursor: pointer; }}
                    .kpi-container {{ display: flex; justify-content: center; gap: 20px; margin-bottom: 40px; }}
                    .kpi-card {{ background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 220px; border-top: 4px solid #0078D7; }}
                    .kpi-card h3 {{ margin: 0; font-size: 32px; color: #0078D7; }}
                    .kpi-card p {{ margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }}
                    .grid-container {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
                    .chart-card {{ background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; display: flex; flex-direction: column; }}
                    table.dataTable thead th {{ background-color: #F0F4F8; color: #333; }}
                </style>
                <script>
                    function switchAI(selectedId) {{
                        const sections = document.querySelectorAll('.dashboard-view-panel');
                        sections.forEach(sec => {{ sec.style.display = 'none'; }});
                        const activeSec = document.getElementById(selectedId);
                        if(activeSec) {{ activeSec.style.display = 'block'; window.dispatchEvent(new Event('resize')); }}
                    }}
                    $(document).ready(function() {{
                        $('.delegate-table').DataTable({{ "order": [[ 3, "desc" ]], "pageLength": 10 }});
                    }});
                </script>
            </head>
            <body>
                <h1>📧 {self.meeting_name} - Mailing List Analytics</h1>
                <div class="selector-container">
                    <label>🎯 Scope / Agenda Item:</label>
                    <select onchange="switchAI(this.value)">{" ".join(dropdown_options)}</select>
                </div>
                {" ".join(views_html_buffer)}
            </body>
            </html>
            """

            out_file = self.export_dir / "Email_Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_template)

            self.finished.emit(True, str(out_file))
        except Exception as e:
            self.finished.emit(False, str(e))

    def _compile_view_block(self, scope_id, total_emails, total_delegates, ai_html, comp_html, time_html, table_html,
                            is_visible=False):
        display = "block" if is_visible else "none"
        ai_card = f'<div class="chart-card">{ai_html}</div>' if ai_html else ""
        col_span = "" if scope_id != "global" else "grid-column: span 1;"

        return f"""
        <div id="{scope_id}" class="dashboard-view-panel" style="display: {display};">
            <div class="kpi-container">
                <div class="kpi-card"><h3>{total_emails}</h3><p>Total Emails</p></div>
                <div class="kpi-card"><h3>{total_delegates}</h3><p>Active Delegates</p></div>
            </div>
            <div class="grid-container">
                {ai_card}
                <div class="chart-card" style="{col_span}">{comp_html}</div>
                <div class="chart-card" style="grid-column: 1 / -1;">{time_html}</div>
                <div class="chart-card" style="grid-column: 1 / -1; overflow-x: auto;">{table_html}</div>
            </div>
        </div>
        """

    def _generate_ai_volume(self, df, prefix):
        valid_df = df[~df['agenda_item'].isin(['UNKNOWN AI', 'UNKNOWN', '', 'NAN', 'NONE'])]
        if valid_df.empty: return ""

        counts = valid_df.groupby('agenda_item').size().reset_index(name='Emails')
        counts.rename(columns={'agenda_item': 'Agenda Item'}, inplace=True)
        counts = counts.sort_values('Emails', ascending=False)

        fig = px.bar(counts, x='Agenda Item', y='Emails', title="Agenda Items by Email Volume",
                     color_discrete_sequence=[self.THEME_COLOR])

        fig.update_xaxes(type='category', categoryorder='total descending')

        # ---> THE FIX: Force a unique div_id so Plotly JS doesn't overwrite earlier charts!
        safe_id = f"ai_vol_{prefix}".replace(" ", "_").replace(".", "_")
        return fig.to_html(full_html=False, include_plotlyjs=False, div_id=safe_id)

    def _generate_company_volume(self, df, prefix):
        valid_df = df[~df['company'].isin(['Unknown', '', 'Nan', 'None'])]
        if valid_df.empty: return ""

        counts = valid_df.groupby('company').size().reset_index(name='Emails')
        counts.rename(columns={'company': 'Company'}, inplace=True)
        counts = counts.sort_values('Emails', ascending=False).head(20)

        fig = px.bar(counts, x='Emails', y='Company', orientation='h', title="Top 20 Active Companies",
                     color_discrete_sequence=[self.THEME_COLOR])

        fig.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1)

        # ---> THE FIX: Unique ID for Company volume
        safe_id = f"comp_vol_{prefix}".replace(" ", "_").replace(".", "_")
        return fig.to_html(full_html=False, include_plotlyjs=False, div_id=safe_id)

    def _generate_timeline(self, df, prefix):
        valid_df = df.dropna(subset=['date_received']).sort_values('date_received').copy()

        if valid_df.empty:
            return "<p style='padding:20px; color:#666;'>No valid temporal data available for this view.</p>"

        fig = px.histogram(
            valid_df,
            x='date_received',
            title="Email Traffic Over Time (1-Hour Bins)",
            color_discrete_sequence=[self.THEME_COLOR]
        )

        fig.update_traces(xbins=dict(size=3600000))

        fig.update_xaxes(
            rangeslider_visible=True,
            title="Timeline",
            tickformat="%a, %b %d<br>%H:%M",
            ticklabelmode="period"
        )

        fig.update_yaxes(title="Email Volume")

        # ---> THE FIX: Unique ID for the Timeline
        safe_id = f"time_vol_{prefix}".replace(" ", "_").replace(".", "_")
        return fig.to_html(full_html=False, include_plotlyjs=False, div_id=safe_id)

    def _generate_delegate_table(self, df, prefix):
        delegates = df.groupby('sender_email').agg(
            Name=('sender_name', lambda x: x.value_counts().index[0] if not x.empty else "Unknown"),
            Company=('company', 'first'),
            Emails=('id', 'count')
        ).reset_index().sort_values('Emails', ascending=False).head(100)

        rows = ""
        for _, row in delegates.iterrows():
            rows += f"<tr><td>{row['Name']}</td><td>{row['sender_email']}</td><td>{row['Company']}</td><td>{row['Emails']}</td></tr>"

        # ---> BEST PRACTICE: Give the table a unique HTML ID as well
        safe_id = f"table_{prefix}".replace(" ", "_").replace(".", "_")
        return f"""
        <h3 style="color:#333; text-align:center;">Top Active Delegates</h3>
        <table id="{safe_id}" class="display delegate-table" style="width:100%">
            <thead><tr><th>Name</th><th>Email Address</th><th>Company</th><th>Sent Emails</th></tr></thead>
            <tbody>{rows}</tbody>
        </table>
        """