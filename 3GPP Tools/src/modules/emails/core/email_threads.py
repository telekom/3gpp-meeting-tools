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

    def __init__(self, source_path: str, meeting_dir: Path, ai_lookup: dict, db: EmailDatabase):
        super().__init__()
        self.source_path = source_path
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup
        self.db = db

    def run(self):
        pythoncom.CoInitialize()
        try:
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

            # ---> THE FIX: Our memory buffer for batching
            batch_data = []

            for i in range(1, total_items + 1):
                mail_item = items.Item(i)
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

                # ---> THE FIX: Flush the batch to SQLite every 50 valid emails
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
        pythoncom.CoInitialize()
        try:
            total = len(self.items_to_move)
            success_count = 0

            # ---> THE FIX: Buffer for DB updates
            batch_updates = []

            for i, (entry_id, ai) in enumerate(self.items_to_move, 1):
                moved = OutlookClient.move_email_to_target(entry_id, self.target_base_path, ai)

                if moved:
                    batch_updates.append(('Target', entry_id))
                    success_count += 1

                # Flush to DB every 20 moves
                if len(batch_updates) >= 20:
                    self.db.update_locations_batch(batch_updates)
                    batch_updates.clear()

                self.progress_update.emit(i, total)

            # Flush the remainder
            if batch_updates:
                self.db.update_locations_batch(batch_updates)

            self.finished.emit(True, f"✅ Successfully moved {success_count}/{total} emails to Target.")
        except Exception as e:
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()


class EmailTargetRescanThread(QThread):
    log_msg = pyqtSignal(str, int)
    progress_update = pyqtSignal(int, int)
    finished = pyqtSignal(bool, str)

    def __init__(self, target_path: str, meeting_dir: Path, ai_lookup: dict, db: EmailDatabase):
        super().__init__()
        self.target_path = target_path
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup
        self.db = db

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            self.log_msg.emit(f"Scanning Target folder: {self.target_path}...", logging.INFO)
            target_base = OutlookClient.get_folder_by_path(self.target_path)

            if not target_base:
                self.finished.emit(False, "Could not find the specified Target folder in Outlook.")
                return

            # We must scan the base target folder AND all the AI subfolders inside it!
            folders_to_scan = [target_base]
            for sub in target_base.Folders:
                folders_to_scan.append(sub)

            total_found = 0
            valid_count = 0
            batch_data = []

            for folder in folders_to_scan:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                total_items = len(items)
                total_found += total_items

                for i in range(1, total_items + 1):
                    mail_item = items.Item(i)
                    if i % 10 == 0: self.progress_update.emit(i, total_found)
                    if mail_item.Class != 43: continue

                    parsed_data = EmailParser.parse_outlook_item(mail_item, self.ai_lookup)

                    if parsed_data and parsed_data.get('tdoc_id'):
                        # ---> DUPLICATE PREVENTION: Check if we already have the .msg file!
                        existing = self.db.get_email(parsed_data['id'])

                        if existing and existing.get('msg_path') and Path(existing['msg_path']).exists():
                            parsed_data['msg_path'] = existing['msg_path']
                        else:
                            parsed_data['msg_path'] = OutlookClient.save_email_to_disk(
                                mail_item, parsed_data['tdoc_id'], self.meeting_dir
                            )

                        # Hardcode the location to Target since we found it here
                        parsed_data['outlook_location'] = 'Target'
                        batch_data.append(parsed_data)
                        valid_count += 1

                        # Flush batch
                        if len(batch_data) >= 50:
                            self.db.save_emails_batch(batch_data)
                            batch_data.clear()

            # Flush remainder
            if batch_data:
                self.db.save_emails_batch(batch_data)

            self.log_msg.emit(f"✅ Rescan complete! Updated {valid_count} emails.", logging.INFO)
            self.finished.emit(True, f"Successfully rescanned {valid_count} Target emails.")

        except Exception as e:
            self.log_msg.emit(f"Error during rescan: {str(e)}", logging.ERROR)
            self.finished.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()