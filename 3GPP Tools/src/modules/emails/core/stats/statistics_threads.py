import re
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

from core.config.plot_styles import THEME_COLOR
from core.utils.company_sanitizer import CompanySanitizer
from modules.emails.core.stats.plot_agenda import _generate_ai_volume
from modules.emails.core.stats.plot_companies import _generate_company_volume, _generate_company_ai_heatmap
from modules.emails.core.stats.plot_delegates import _generate_delegate_table, _generate_delegates_plot
from modules.emails.core.stats.plot_timeline import _generate_timeline


class EmailStatsExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    # ---> THE FIX: Added 'config: dict' to the signature
    def __init__(self, meeting_dir: Path, email_data: list, meeting_name: str, config: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.email_data = email_data
        self.meeting_name = meeting_name
        self.config = config
        self.export_dir = self.meeting_dir / "Export"
        self.THEME_COLOR = THEME_COLOR

        # Extract dynamic settings loaded from email_config.json
        self.cfg_top_comps = self.config.get("email_top_companies", 25)
        self.cfg_top_dels = self.config.get("email_top_delegates", 25)
        self.cfg_hm_comps = self.config.get("email_heatmap_top_comps", 25)
        self.cfg_hm_ais = self.config.get("email_heatmap_top_ais", 25)

        self.svg_config = {
            'toImageButtonOptions': {
                'format': 'svg',
                'filename': 'email_statistics_plot'
            }
        }

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)
            df = pd.DataFrame(self.email_data)

            if df.empty:
                self.finished.emit(False, "No email data available to generate statistics.")
                return

            df['date_received'] = pd.to_datetime(df['date_received'], utc=True, errors='coerce').dt.tz_localize(None)

            # =================================================================
            # 🧹 DATA HEALING
            # =================================================================
            df['agenda_item'] = df['agenda_item'].astype(str).str.strip().str.upper()
            df['company'] = df['company'].astype(str).str.strip().str.title()
            df['sender_email'] = df['sender_email'].astype(str).str.strip().str.lower()

            def unify_company(row):
                raw_str = f"{row.get('sender_name', '')} <{row.get('sender_email', '')}>"
                matches = CompanySanitizer.get_matching_contributors(raw_str)
                if matches:
                    return matches
                comp = str(row.get('company', '')).strip().title()
                return [comp] if comp and comp not in ['None', 'Nan', ''] else []

            df['Clean_Companies'] = df.apply(unify_company, axis=1)

            def split_ais(ai_str):
                ai_str = str(ai_str).upper().replace('AND', ',').replace('&', ',')
                return [ai.strip() for ai in re.split(r'[,/]', ai_str) if
                        ai.strip() and ai.strip() not in ['UNKNOWN AI', 'UNKNOWN', 'NAN', 'NONE']]

            df['ai_list'] = df['agenda_item'].apply(split_ais)

            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

            all_ais = set([ai for sublist in df['ai_list'] for ai in sublist])
            unique_ais = sorted(list(all_ais), key=natural_sort_key)

            # =================================================================
            # 📊 GENERATE GLOBAL VIEW
            # =================================================================
            g_html_ai = _generate_ai_volume(self.THEME_COLOR, self.svg_config, df, "Global", include_plotlyjs='cdn')

            # Pass configurations to the functions
            g_html_comp = _generate_company_volume(self.THEME_COLOR, self.svg_config, df, "Global", False,
                                                   self.cfg_top_comps)
            g_html_dels = _generate_delegates_plot(self.THEME_COLOR, self.svg_config, df, "Global", False,
                                                   self.cfg_top_dels)
            g_html_hm = _generate_company_ai_heatmap(self.svg_config, df, "Global", False, self.cfg_hm_comps,
                                                     self.cfg_hm_ais)

            g_html_time = _generate_timeline(self.THEME_COLOR, self.svg_config, df, "Global", False)
            g_html_table = _generate_delegate_table(df, "Global")

            views_html_buffer = []
            dropdown_options = ['<option value="global">🌐 Overall Email Analytics</option>']

            views_html_buffer.append(self._compile_view_block(
                "global", len(df), df['sender_email'].nunique(),
                g_html_ai, g_html_comp, g_html_dels, g_html_hm, g_html_time, g_html_table, is_visible=True
            ))

            # =================================================================
            # 📊 GENERATE PER-AI VIEWS
            # =================================================================
            for idx, ai_name in enumerate(unique_ais):
                ai_df = df[df['ai_list'].apply(lambda x: ai_name in x)].copy()
                if ai_df.empty: continue

                safe_id = f"ai_{idx}"
                dropdown_options.append(
                    f'<option value="{safe_id}">📌 Agenda Item {ai_name} ({len(ai_df)} Emails)</option>')

                ai_html_comp = _generate_company_volume(self.THEME_COLOR, self.svg_config, ai_df, safe_id, False,
                                                        self.cfg_top_comps)
                ai_html_dels = _generate_delegates_plot(self.THEME_COLOR, self.svg_config, ai_df, safe_id, False,
                                                        self.cfg_top_dels)
                ai_html_time = _generate_timeline(self.THEME_COLOR, self.svg_config, ai_df, safe_id, False)
                ai_html_table = _generate_delegate_table(ai_df, safe_id)

                views_html_buffer.append(self._compile_view_block(
                    safe_id, len(ai_df), ai_df['sender_email'].nunique(),
                    None, ai_html_comp, ai_html_dels, None, ai_html_time, ai_html_table, is_visible=False
                ))

            # =================================================================
            # 📄 ASSEMBLE HTML
            # =================================================================
            dashboard_template = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>📧 Email Analytics - __MEETING_NAME__</title>
                <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
                <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
                <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
                <style>
                    body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAFAFA; margin: 0; padding: 20px; }
                    h1 { color: #333; text-align: center; margin-bottom: 10px; }
                    .selector-container { display: flex; justify-content: center; margin-bottom: 30px; background: #FFF; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #E0E0E0; }
                    .selector-container label { font-weight: bold; margin-right: 12px; align-self: center; color: #444; }
                    select { padding: 8px 16px; border-radius: 6px; border: 1px solid #CCCCCC; font-size: 14px; font-weight: bold; color: #005A9E; outline: none; background: #F4F8FC; cursor: pointer; }
                    .kpi-container { display: flex; justify-content: center; gap: 20px; margin-bottom: 40px; }
                    .kpi-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; text-align: center; width: 220px; border-top: 4px solid #0078D7; }
                    .kpi-card h3 { margin: 0; font-size: 32px; color: #0078D7; }
                    .kpi-card p { margin: 5px 0 0; color: #666; font-size: 14px; text-transform: uppercase; font-weight: bold; }
                    .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
                    .chart-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 20px; display: flex; flex-direction: column; height: 100%; min-height: 400px; }
                    table.dataTable thead th { background-color: #F0F4F8; color: #333; }
                </style>
                <script>
                    function switchAI(selectedId) {
                        const sections = document.querySelectorAll('.dashboard-view-panel');
                        sections.forEach(sec => { sec.style.display = 'none'; });
                        const activeSec = document.getElementById(selectedId);
                        if(activeSec) { 
                            activeSec.style.display = 'block'; 
                            window.dispatchEvent(new Event('resize')); 
                        }
                    }
                    $(document).ready(function() {
                        $('.delegate-table').DataTable({ "order": [[ 3, "desc" ]], "pageLength": 10 });
                    });
                </script>
            </head>
            <body>
                <h1>📧 __MEETING_NAME__ - Mailing List Analytics</h1>
                <div class="selector-container">
                    <label>🎯 Scope / Agenda Item:</label>
                    <select onchange="switchAI(this.value)">
                        __DROPDOWN_OPTIONS__
                    </select>
                </div>
                __VIEWS_HTML__
            </body>
            </html>
            """

            dashboard_template = dashboard_template.replace("__MEETING_NAME__", str(self.meeting_name))
            dashboard_template = dashboard_template.replace("__DROPDOWN_OPTIONS__", " ".join(dropdown_options))
            dashboard_template = dashboard_template.replace("__VIEWS_HTML__", " ".join(views_html_buffer))

            out_file = self.export_dir / "Email_Statistics_Report.html"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(dashboard_template)

            self.finished.emit(True, str(out_file))
        except Exception as e:
            self.finished.emit(False, str(e))

    def _compile_view_block(self, scope_id, total_emails, total_delegates, ai_html, comp_html, dels_html, hm_html, time_html, table_html, is_visible=False):
        display_style = "block" if is_visible else "none"

        ai_card = ""
        if ai_html:
            ai_card = f'<div class="chart-card" style="grid-column: 1 / -1;">{ai_html}</div>'

        # ---> THE FIX 3: Increased container height to 850px and added the Expansion button/Tooltip
        hm_card = ""
        if hm_html:
            hm_card = f"""
            <div class="chart-card" style="grid-column: 1 / -1; height: 850px;">
                <div class="info-title-container">
                    <span class="custom-tooltip">ⓘ<span class="custom-tooltip-text"><b>Company Focus Matrix:</b> Heatmap of email traffic by the top companies across top topics.</span></span>
                </div>
                <button class="fs-btn" onclick="toggleFullscreen(this)">⛶ Expand</button>
                {hm_html}
            </div>
            """

        html_template = """
        <div id="__SCOPE_ID__" class="dashboard-view-panel" style="display: __DISPLAY_STYLE__;">
            <div class="kpi-container">
                <div class="kpi-card"><h3>__TOTAL_EMAILS__</h3><p>Total Emails</p></div>
                <div class="kpi-card"><h3>__TOTAL_DELEGATES__</h3><p>Active Delegates</p></div>
            </div>
            <div class="grid-container">
                __AI_CARD__
                <div class="chart-card" style="grid-column: span 1;">
                    __COMP_HTML__
                </div>
                <div class="chart-card" style="grid-column: span 1;">
                    __DELS_HTML__
                </div>
                __HM_CARD__
                <div class="chart-card" style="grid-column: 1 / -1;">
                    __TIME_HTML__
                </div>
                <div class="chart-card" style="grid-column: 1 / -1; overflow-x: auto;">
                    __TABLE_HTML__
                </div>
            </div>
        </div>
        """

        html_template = html_template.replace("__SCOPE_ID__", str(scope_id))
        html_template = html_template.replace("__DISPLAY_STYLE__", str(display_style))
        html_template = html_template.replace("__TOTAL_EMAILS__", str(total_emails))
        html_template = html_template.replace("__TOTAL_DELEGATES__", str(total_delegates))
        html_template = html_template.replace("__AI_CARD__", str(ai_card))
        html_template = html_template.replace("__COMP_HTML__", str(comp_html))
        html_template = html_template.replace("__DELS_HTML__", str(dels_html))
        html_template = html_template.replace("__HM_CARD__", str(hm_card))
        html_template = html_template.replace("__TIME_HTML__", str(time_html))
        html_template = html_template.replace("__TABLE_HTML__", str(table_html))

        return html_template