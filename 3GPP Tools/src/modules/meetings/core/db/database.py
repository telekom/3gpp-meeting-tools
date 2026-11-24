"""MeetingsDatabase orchestrator combining schema/writes and query mixins."""

from __future__ import annotations

from pathlib import Path

from modules.meetings.core.db.queries import MeetingsQueriesMixin
from modules.meetings.core.db.schema_and_writes import MeetingsSchemaAndWritesMixin


class MeetingsDatabase(MeetingsSchemaAndWritesMixin, MeetingsQueriesMixin):
    """SQLite persistence layer for 3GPP working groups and meeting catalogues."""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._create_tables()
