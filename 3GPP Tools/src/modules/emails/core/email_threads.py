# --- File: modules/emails/core/email_threads.py ---
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.email_parser import EmailParser
from modules.emails.core.email_db import EmailDatabase
import logging
import pythoncom


class EmailSyncThread(QThread):
    # Signals to update the UI safely
    log_msg = pyqtSignal(str, int)
    progress_update = pyqtSignal(int, int)  # (current, total)
    finished = pyqtSignal(bool, str)

    # ---> FIX: Added start_date and end_date to the parameters!
    def __init__(self, source_path: str, meeting_dir: Path, ai_lookup: dict, db: EmailDatabase, start_date: str = "",
                 end_date: str = ""):
        super().__init__()
        self.source_path = source_path
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup
        self.db = db
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        import pythoncom
        import datetime
        pythoncom.CoInitialize()
        try:
            # ---> NEW: Parse Dates and apply +/- 3 day buffer
            filter_start, filter_end = None, None
            if self.start_date and self.end_date:
                start_dt = datetime.datetime.strptime(self.start_date, "%Y-%m-%d")
                end_dt = datetime.datetime.strptime(self.end_date, "%Y-%m-%d")
                filter_start = start_dt - datetime.timedelta(days=3)
                filter_end = end_dt + datetime.timedelta(days=4)  # +4 ensures we cover the end of the final day

            self.log_msg.emit(f"Connecting to Outlook folder: {self.source_path}...", logging.INFO)
            source_folder = OutlookClient.get_folder_by_path(self.source_path)

            if not source_folder:
                self.finished.emit(False, "Could not find the specified Source Outlook folder.")
                return

            items = source_folder.Items
            total_items = len(items)
            self.log_msg.emit(f"Found {total_items} items. Scanning for 3GPP eMeeting emails...", logging.INFO)

            items.Sort("[ReceivedTime]", True)

            processed_count = 0
            valid_count = 0
            batch_data = []

            for i in range(1, total_items + 1):
                mail_item = items.Item(i)

                # ---> NEW: Date Range Enforcement logic
                if filter_start and filter_end:
                    mail_date = getattr(mail_item, "ReceivedTime", None)
                    if mail_date:
                        try:
                            # Strip out pywintypes timezone data to create a naive comparable datetime
                            dt = datetime.datetime(mail_date.year, mail_date.month, mail_date.day,
                                                   mail_date.hour, mail_date.minute, mail_date.second)
                            if dt > filter_end:
                                continue  # Email arrived after the meeting ended, skip to next.
                            if dt < filter_start:
                                # FAST EXIT: Because we sorted Newest->Oldest, if this email is older
                                # than our start buffer, ALL remaining emails are even older! Terminate loop!
                                break
                        except Exception:
                            pass

                if i % 10 == 0: self.progress_update.emit(i, total_items)
                if mail_item.Class != 43: continue

                parsed_data = EmailParser.parse_outlook_item(mail_item, self.ai_lookup)

                if parsed_data and parsed_data.get('tdoc_id'):
                    msg_path = OutlookClient.save_email_to_disk(mail_item, parsed_data['tdoc_id'], self.meeting_dir)
                    parsed_data['msg_path'] = msg_path
                    parsed_data['outlook_location'] = 'Source'

                    # Add to our buffer instead of hitting the database directly
                    batch_data.append(parsed_data)
                    valid_count += 1

                # Flush the batch to SQLite every 50 valid emails
                if len(batch_data) >= 50:
                    self.db.save_emails_batch(batch_data)
                    batch_data.clear()

                processed_count += 1

            # Flush any remaining emails in the buffer at the end
            if batch_data:
                self.db.save_emails_batch(batch_data)

            self.progress_update.emit(total_items, total_items)
            self.log_msg.emit(f"✅ Sync complete! Extracted {valid_count} valid TDoc emails.", logging.INFO)
            self.finished.emit(True, f"Successfully synced {valid_count} emails.")

        except Exception as e:
            self.log_msg.emit(f"Fatal error during sync: {str(e)}", logging.ERROR)
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()


class EmailMoveThread(QThread):
    progress_update = pyqtSignal(int, int)
    finished = pyqtSignal(bool, str)

    def __init__(self, items_to_move: list, target_base_path: str, db: EmailDatabase):
        super().__init__()
        # items_to_move is a list of tuples: [(entry_id, agenda_item), ...]
        self.items_to_move = items_to_move
        self.target_base_path = target_base_path
        self.db = db

    def run(self):
        import pythoncom
        import sqlite3  # <--- Ensure sqlite3 is imported for the cleanup query
        pythoncom.CoInitialize()
        try:
            total = len(self.items_to_move)
            success_count = 0
            ghost_count = 0

            # Buffer for DB updates
            batch_updates = []

            for i, (entry_id, ai) in enumerate(self.items_to_move, 1):
                status = OutlookClient.move_email_to_target(entry_id, self.target_base_path, ai)

                if status == "SUCCESS" or status is True:
                    batch_updates.append(('Target', entry_id))
                    success_count += 1
                elif status == "DELETED":
                    # ---> SELF-HEALING: Purge the deleted email from the local database
                    with sqlite3.connect(self.db.db_path) as conn:
                        conn.execute('DELETE FROM emails WHERE id = ?', (entry_id,))
                        conn.commit()
                    ghost_count += 1

                # Flush to DB every 20 moves
                if len(batch_updates) >= 20:
                    self.db.update_locations_batch(batch_updates)
                    batch_updates.clear()

                self.progress_update.emit(i, total)

            # Flush the remainder
            if batch_updates:
                self.db.update_locations_batch(batch_updates)

            msg = f"✅ Successfully moved {success_count}/{total} emails to Target."
            if ghost_count > 0:
                msg += f" (Cleaned up {ghost_count} deleted emails from database)."

            self.finished.emit(True, msg)
        except Exception as e:
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()


class EmailTargetRescanThread(QThread):
    log_msg = pyqtSignal(str, int)
    progress_update = pyqtSignal(int, int)
    finished = pyqtSignal(bool, str)

    # ---> FIX: Added start_date and end_date to the parameters!
    def __init__(self, target_path: str, meeting_dir: Path, ai_lookup: dict, db: EmailDatabase, start_date: str = "",
                 end_date: str = ""):
        super().__init__()
        self.target_path = target_path
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup
        self.db = db
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            # ---> NEW: Parse Dates and apply +/- 3 day buffer
            filter_start, filter_end = None, None
            if self.start_date and self.end_date:
                import datetime
                start_dt = datetime.datetime.strptime(self.start_date, "%Y-%m-%d")
                end_dt = datetime.datetime.strptime(self.end_date, "%Y-%m-%d")
                filter_start = start_dt - datetime.timedelta(days=3)
                filter_end = end_dt + datetime.timedelta(days=4)  # +4 ensures we cover the end of the final day

            self.log_msg.emit(f"Scanning Target folder: {self.target_path}...", logging.INFO)
            target_base = OutlookClient.get_folder_by_path(self.target_path)

            if not target_base:
                self.finished.emit(False, "Could not find the specified Target folder in Outlook.")
                return

            folders_to_scan = [target_base]
            for sub in target_base.Folders:
                folders_to_scan.append(sub)

            total_items_to_scan = 0
            for folder in folders_to_scan:
                total_items_to_scan += len(folder.Items)

            self.log_msg.emit(f"Found {total_items_to_scan} total items. Scanning...", logging.INFO)

            processed_count = 0
            valid_count = 0
            batch_data = []

            for folder in folders_to_scan:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)  # Sort newest first!
                total_in_folder = len(items)

                for i in range(1, total_in_folder + 1):
                    processed_count += 1

                    if processed_count % 10 == 0:
                        self.progress_update.emit(processed_count, total_items_to_scan)

                    mail_item = items.Item(i)

                    # ---> NEW: Date Range Enforcement logic
                    if filter_start and filter_end:
                        mail_date = getattr(mail_item, "ReceivedTime", None)
                        if mail_date:
                            try:
                                dt = datetime.datetime(mail_date.year, mail_date.month, mail_date.day,
                                                       mail_date.hour, mail_date.minute, mail_date.second)
                                if dt > filter_end:
                                    continue  # Skip future/newer emails
                                if dt < filter_start:
                                    break  # FAST EXIT: Stop scanning this specific subfolder!
                            except Exception:
                                pass

                    if mail_item.Class != 43: continue

                    parsed_data = EmailParser.parse_outlook_item(mail_item, self.ai_lookup)

                    if parsed_data and parsed_data.get('tdoc_id'):
                        # ---> NEW: Check if email is already in the database
                        existing = self.db.get_email(parsed_data['id'])

                        if existing and existing.get('msg_path') and Path(existing['msg_path']).exists():
                            # Skip saving to disk, reuse existing path
                            parsed_data['msg_path'] = existing['msg_path']
                        else:
                            # Save new .msg file to disk
                            msg_path = OutlookClient.save_email_to_disk(mail_item, parsed_data['tdoc_id'],
                                                                        self.meeting_dir)
                            parsed_data['msg_path'] = msg_path

                        parsed_data['outlook_location'] = 'Source'

                        # Add to our buffer. The DB's "INSERT OR REPLACE" will seamlessly
                        # update the sender and company fields for existing emails!
                        batch_data.append(parsed_data)
                        valid_count += 1

                        if len(batch_data) >= 50:
                            self.db.save_emails_batch(batch_data)
                            batch_data.clear()

            if batch_data:
                self.db.save_emails_batch(batch_data)

            self.progress_update.emit(total_items_to_scan, total_items_to_scan)
            self.log_msg.emit(f"✅ Rescan complete! Updated {valid_count} emails.", logging.INFO)
            self.finished.emit(True, f"Successfully rescanned {valid_count} Target emails.")

        except Exception as e:
            self.log_msg.emit(f"Error during rescan: {str(e)}", logging.ERROR)
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()


import pandas as pd
import plotly.express as px
import re


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

            # ---> FIX 1: Parse dates, but DO NOT drop missing rows from the main dataframe!
            # This ensures emails with weird dates still count towards AI and Company totals.
            df['date_received'] = pd.to_datetime(df['date_received'], utc=True, errors='coerce').dt.tz_localize(None)

            # Strip whitespace to prevent fragmentation
            df['agenda_item'] = df['agenda_item'].astype(str).str.strip()
            df['company'] = df['company'].astype(str).str.strip()

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
        # Explicitly exclude 'Unknown AI' and blanks
        valid_df = df[~df['agenda_item'].isin(['Unknown AI', 'Unknown', '', 'nan', 'None'])]
        counts = valid_df['agenda_item'].value_counts().reset_index().head(20)
        counts.columns = ['Agenda Item', 'Emails']

        fig = px.bar(counts, x='Agenda Item', y='Emails', title="Top 20 Agenda Items by Email Volume",
                     color_discrete_sequence=[self.THEME_COLOR])

        # ---> THE MAGIC FIX: Force Plotly to treat AIs as text, preventing the "Date" conversion bug!
        fig.update_xaxes(type='category', categoryorder='total descending')

        return fig.to_html(full_html=False, include_plotlyjs=False)

    def _generate_company_volume(self, df, prefix):
        # Exclude blanks or 'Unknown' companies
        valid_df = df[~df['company'].isin(['Unknown', '', 'nan', 'None'])]
        counts = valid_df['company'].value_counts().reset_index().head(20)
        counts.columns = ['Company', 'Emails']

        fig = px.bar(counts, x='Emails', y='Company', orientation='h', title="Top 20 Active Companies",
                     color_discrete_sequence=[self.THEME_COLOR])

        # ---> BEST PRACTICE: Force categorical type on the Y-axis and let Plotly handle the sorting!
        fig.update_yaxes(
            type='category',
            categoryorder='total ascending',  # 'ascending' puts the biggest bar at the top for horizontal charts
            tickmode='linear',
            dtick=1
        )

        return fig.to_html(full_html=False, include_plotlyjs=False)

    def _generate_timeline(self, df, prefix):
        # ---> FIX 2: Only drop missing dates locally for the timeline chart
        valid_df = df.dropna(subset=['date_received']).sort_values('date_received').copy()

        if valid_df.empty:
            return "<p style='padding:20px; color:#666;'>No valid temporal data available for this view.</p>"

        fig = px.histogram(
            valid_df,
            x='date_received',
            title="Email Traffic Over Time (1-Hour Bins)",
            color_discrete_sequence=[self.THEME_COLOR]
        )

        # ---> FIX 3: Force the bins to exactly 1 hour (3,600,000 milliseconds)
        fig.update_traces(xbins=dict(size=3600000))

        # Add the Day of the Week (%a) to the x-axis ticks
        fig.update_xaxes(
            rangeslider_visible=True,
            title="Timeline",
            tickformat="%a, %b %d<br>%H:%M",
            ticklabelmode="period"
        )

        fig.update_yaxes(title="Email Volume")

        return fig.to_html(full_html=False, include_plotlyjs=False)

    def _generate_delegate_table(self, df, prefix):
        # Get the most common display name per email address to bypass listserv aliases
        delegates = df.groupby('sender_email').agg(
            Name=('sender_name', lambda x: x.value_counts().index[0] if not x.empty else "Unknown"),
            Company=('company', 'first'),
            Emails=('id', 'count')
        ).reset_index().sort_values('Emails', ascending=False).head(100)  # Top 100 for performance

        rows = ""
        for _, row in delegates.iterrows():
            rows += f"<tr><td>{row['Name']}</td><td>{row['sender_email']}</td><td>{row['Company']}</td><td>{row['Emails']}</td></tr>"

        return f"""
        <h3 style="color:#333; text-align:center;">Top Active Delegates</h3>
        <table class="display delegate-table" style="width:100%">
            <thead><tr><th>Name</th><th>Email Address</th><th>Company</th><th>Sent Emails</th></tr></thead>
            <tbody>{rows}</tbody>
        </table>
        """