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
        # ---> NEU: Initialisiert das Windows COM-System für diesen Hintergrund-Thread
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

            processed_count = 0
            valid_count = 0

            # Sort items by received time (descending) so we process newest first
            items.Sort("[ReceivedTime]", True)

            for i in range(1, total_items + 1):
                mail_item = items.Item(i)

                # Update progress bar every 10 items
                if i % 10 == 0:
                    self.progress_update.emit(i, total_items)

                # Skip non-mail items (like meeting invites or delivery receipts)
                if mail_item.Class != 43:  # 43 = olMail
                    continue

                # 1. Parse the email
                parsed_data = EmailParser.parse_outlook_item(mail_item, self.ai_lookup)

                # If it's an eMeeting email (returned valid data)
                if parsed_data and parsed_data.get('tdoc_id'):
                    # 2. Save physical .msg file
                    msg_path = OutlookClient.save_email_to_disk(
                        mail_item,
                        parsed_data['tdoc_id'],
                        self.meeting_dir
                    )
                    parsed_data['msg_path'] = msg_path

                    # 3. Save metadata to SQLite
                    parsed_data['outlook_location'] = 'Source'
                    self.db.save_email(parsed_data)
                    valid_count += 1

                processed_count += 1

            self.progress_update.emit(total_items, total_items)
            self.log_msg.emit(f"✅ Sync complete! Extracted {valid_count} valid TDoc emails.", logging.INFO)
            self.finished.emit(True, f"Successfully synced {valid_count} emails.")

        except Exception as e:
            self.log_msg.emit(f"Fatal error during sync: {str(e)}", logging.ERROR)
            self.finished.emit(False, str(e))
        finally:
            # ---> NEU: Gibt die COM-Ressourcen sauber wieder frei
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
        # ---> NEU: Initialisiert das Windows COM-System auch für den Move-Thread
        pythoncom.CoInitialize()
        try:
            total = len(self.items_to_move)
            success_count = 0

            for i, (entry_id, ai) in enumerate(self.items_to_move, 1):
                # Execute the Outlook COM move
                moved = OutlookClient.move_email_to_target(entry_id, self.target_base_path, ai)

                if moved:
                    # Update our SQLite tracker to reflect the new reality
                    self.db.update_location(entry_id, 'Target')
                    success_count += 1

                self.progress_update.emit(i, total)

            self.finished.emit(True, f"✅ Successfully moved {success_count}/{total} emails to Target.")
        except Exception as e:
            self.finished.emit(False, str(e))
        finally:
            # ---> NEU: Ressourcen freigeben
            pythoncom.CoUninitialize()