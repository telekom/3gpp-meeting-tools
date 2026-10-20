# --- File: src/modules/nas/ui/nas_workers.py ---
import logging
from typing import List, Optional

from PyQt5.QtCore import QThread, pyqtSignal

from modules.nas.core.nas_db import NASDatabase

logger = logging.getLogger(__name__)


class NASMessageListWorker(QThread):
    """Background worker to query message or IE lists without freezing the UI."""

    results_ready = pyqtSignal(list, int)  # (items, request_id)

    def __init__(
            self,
            db: NASDatabase,
            ie_query: str,
            version_ids: List[int],
            search_desc: bool,
            request_id: int,
            browse_mode: str = "messages",
    ):
        super().__init__()
        self.db = db
        self.ie_query = ie_query
        self.version_ids = version_ids
        self.search_desc = search_desc
        self.request_id = request_id
        self.browse_mode = browse_mode

    def run(self):
        try:
            if self.browse_mode == "ies":
                items = self.db.get_ie_definitions_list(
                    version_ids=self.version_ids,
                    ie_query=self.ie_query,
                    search_descriptions=self.search_desc,
                )
            else:
                if self.ie_query:
                    items = self.db.get_messages_by_ie_search(
                        ie_query=self.ie_query,
                        version_ids=self.version_ids,
                        search_descriptions=self.search_desc,
                    )
                else:
                    items = self.db.get_messages_list(self.version_ids)
            self.results_ready.emit(items, self.request_id)
        except Exception as e:
            logger.error(f"Error querying protocol items in background: {e}")
            self.results_ready.emit([], self.request_id)


class ReverseLookupWorker(QThread):
    """Background worker to fetch message references without freezing the UI."""

    results_ready = pyqtSignal(list, int)

    def __init__(
            self,
            db: NASDatabase,
            clause: str,
            ie_name: str,
            spec_number: Optional[str],
            version_ids: List[int],
            request_id: int,
    ):
        super().__init__()
        self.db = db
        self.clause = clause
        self.ie_name = ie_name
        self.spec_number = spec_number
        self.version_ids = version_ids
        self.request_id = request_id

    def run(self):
        try:
            results = self.db.get_messages_using_ie(
                clause=self.clause,
                ie_name=self.ie_name,
                spec_number=self.spec_number,
                version_ids=self.version_ids,
            )
            self.results_ready.emit(results, self.request_id)
        except Exception:
            self.results_ready.emit([], self.request_id)
