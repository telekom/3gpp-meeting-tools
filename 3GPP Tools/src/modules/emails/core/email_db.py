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
        """Initializes the schema for the per-meeting email database."""
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
            # Create an index to make AI and TDoc filtering lightning fast for the UI
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_tdoc ON emails(tdoc_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_ai ON emails(agenda_item)')
            conn.commit()

    def save_email(self, email_data: dict):
        """Inserts or overwrites an email in the local database."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO emails 
                (id, tdoc_id, agenda_item, sender_name, company, date_received, subject, revisions_mentioned, short_text, free_text, msg_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                email_data.get('id'),
                email_data.get('tdoc_id'),
                email_data.get('agenda_item'),
                email_data.get('sender_name'),
                email_data.get('company'),
                email_data.get('date_received'),
                email_data.get('subject'),
                email_data.get('revisions_mentioned'),
                email_data.get('short_text'),
                email_data.get('free_text'),
                email_data.get('msg_path')
            ))
            conn.commit()

    def get_emails_for_tdoc(self, tdoc_id: str) -> list:
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM emails WHERE tdoc_id = ? ORDER BY date_received DESC', (tdoc_id,))
            return [dict(row) for row in cursor.fetchall()]