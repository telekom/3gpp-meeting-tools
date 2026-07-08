# --- File: modules/emails/core/email_threads.py ---
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.email_parser import EmailParser
from modules.emails.core.email_db import EmailDatabase
import logging


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


