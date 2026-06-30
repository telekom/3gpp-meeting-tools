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