# --- File: src/modules/nas/core/db/database.py ---
"""Composed NASDatabase class backed by 3gpp_protocol_data.db."""

import logging
from pathlib import Path

from .ingestion import NasDbIngestionMixin
from .queries_ies import NasDbIeQueriesMixin
from .queries_messages import NasDbMessageQueriesMixin
from .schema import NasDbSchemaMixin


class NASDatabase(
    NasDbSchemaMixin,
    NasDbIngestionMixin,
    NasDbMessageQueriesMixin,
    NasDbIeQueriesMixin,
):
    """Manages the SQLite database for 3GPP NAS (24.501, 24.301) and ASN.1 (38.331, 36.331, 38.413) protocols."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self._init_db()
