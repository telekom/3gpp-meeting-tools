# --- File: modules/emails/core/email_db.py ---
import sqlite3
import logging
from pathlib import Path

class EmailDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS emails (
                    id TEXT PRIMARY KEY,
                    tdoc_id TEXT,
                    agenda_item TEXT,
                    sender_name TEXT,
                    company TEXT,
                    date_received TEXT,
                    subject TEXT,
                    revisions_mentioned TEXT,
                    short_text TEXT,
                    free_text TEXT,
                    msg_path TEXT
                )
            ''')
            cursor.execute('CREATE TABLE IF NOT EXISTS starred_tdocs (tdoc_id TEXT PRIMARY KEY)')
            # Safe Schema Upgrade: Add outlook_location if it doesn't exist
            cursor.execute("PRAGMA table_info(emails)")
            columns = [info[1] for info in cursor.fetchall()]
            if 'outlook_location' not in columns:
                cursor.execute("ALTER TABLE emails ADD COLUMN outlook_location TEXT DEFAULT 'Source'")

            cursor.execute('CREATE INDEX IF NOT EXISTS idx_tdoc ON emails(tdoc_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_ai ON emails(agenda_item)')
            conn.commit()

    def save_email(self, email_data: dict):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO emails 
                (id, tdoc_id, agenda_item, sender_name, company, date_received, subject, 
                 revisions_mentioned, short_text, free_text, msg_path, outlook_location)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                email_data.get('id'), email_data.get('tdoc_id'), email_data.get('agenda_item'),
                email_data.get('sender_name'), email_data.get('company'), email_data.get('date_received'),
                email_data.get('subject'), email_data.get('revisions_mentioned'), email_data.get('short_text'),
                email_data.get('free_text'), email_data.get('msg_path'), email_data.get('outlook_location', 'Source')
            ))
            conn.commit()

    def update_location(self, entry_id: str, new_location: str):
        """Updates the tracked location after an explicit move."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE emails SET outlook_location = ? WHERE id = ?', (new_location, entry_id))
            conn.commit()

    def save_emails_batch(self, emails_data: list):
        """Mass inserts/updates a list of emails in a single blazing-fast transaction."""
        if not emails_data: return
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # Convert the list of dictionaries into a list of tuples for executemany
            tuples = [
                (e.get('id'), e.get('tdoc_id'), e.get('agenda_item'),
                 e.get('sender_name'), e.get('company'), e.get('date_received'),
                 e.get('subject'), e.get('revisions_mentioned'), e.get('short_text'),
                 e.get('free_text'), e.get('msg_path'), e.get('outlook_location', 'Source'))
                for e in emails_data
            ]

            cursor.executemany('''
                INSERT OR REPLACE INTO emails 
                (id, tdoc_id, agenda_item, sender_name, company, date_received, subject, 
                 revisions_mentioned, short_text, free_text, msg_path, outlook_location)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', tuples)
            conn.commit()

    def update_locations_batch(self, location_updates: list):
        """Mass updates Outlook locations. Expects a list of tuples: [('Target', 'entryID_123'), ...]"""
        if not location_updates: return
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.executemany('UPDATE emails SET outlook_location = ? WHERE id = ?', location_updates)
            conn.commit()

    def save_emails_batch(self, emails_data: list):
        """Mass inserts/updates a list of emails in a single blazing-fast transaction."""
        if not emails_data: return
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # Convert the list of dictionaries into a list of tuples for executemany
            tuples = [
                (e.get('id'), e.get('tdoc_id'), e.get('agenda_item'),
                 e.get('sender_name'), e.get('company'), e.get('date_received'),
                 e.get('subject'), e.get('revisions_mentioned'), e.get('short_text'),
                 e.get('free_text'), e.get('msg_path'), e.get('outlook_location', 'Source'))
                for e in emails_data
            ]

            cursor.executemany('''
                INSERT OR REPLACE INTO emails 
                (id, tdoc_id, agenda_item, sender_name, company, date_received, subject, 
                 revisions_mentioned, short_text, free_text, msg_path, outlook_location)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', tuples)
            conn.commit()

    def update_locations_batch(self, location_updates: list):
        """Mass updates Outlook locations. Expects a list of tuples: [('Target', 'entryID_123'), ...]"""
        if not location_updates: return
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.executemany('UPDATE emails SET outlook_location = ? WHERE id = ?', location_updates)
            conn.commit()

    def get_email(self, entry_id: str) -> dict:
        """Fetches a single email by its Outlook EntryID."""
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM emails WHERE id = ?', (entry_id,))
            row = cursor.fetchone()
            return dict(row) if row else {}

    def toggle_tdoc_star(self, tdoc_id: str, star: bool):
        """Marks or unmarks a TDoc as starred."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if star:
                cursor.execute('INSERT OR IGNORE INTO starred_tdocs (tdoc_id) VALUES (?)', (tdoc_id,))
            else:
                cursor.execute('DELETE FROM starred_tdocs WHERE tdoc_id = ?', (tdoc_id,))
            conn.commit()

    def get_starred_tdocs(self) -> set:
        """Returns a set of all currently starred TDoc IDs."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT tdoc_id FROM starred_tdocs')
            return {row[0] for row in cursor.fetchall()}