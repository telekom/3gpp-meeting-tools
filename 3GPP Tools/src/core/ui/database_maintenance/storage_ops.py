"""Pure SQLite storage discovery, metrics querying, and compaction operations."""

from __future__ import annotations

import logging
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)


def format_file_size(size_bytes: int) -> str:
    """Formats raw byte count into a human-readable string (KB, MB, GB)."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    for unit in ["KB", "MB", "GB"]:
        size_bytes /= 1024.0
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
    return f"{size_bytes:.2f} TB"


def get_tracked_databases(root: Optional[Path] = None) -> List[Dict[str, Any]]:
    """Identifies all standard databases in the application root and subdirectories."""
    base_root = root if root is not None else get_project_root()
    candidate_files = [
        {
            "name": "3GPP Core Data (Specs, Meetings, Work Items)",
            "subtitle": "Specifications, meetings, and work items",
            "icon": "📋",
            "path": base_root / "db" / "3gpp_data.db",
        },
        {
            "name": "3GPP Protocol Data (NAS, NGAP, RRC, etc.)",
            "subtitle": "Protocol messages and IE definitions",
            "icon": "📡",
            "path": base_root / "db" / "3gpp_protocol_data.db",
        },
        {
            "name": "3GPP Specification Text Search Index",
            "subtitle": "FTS5 trigram clause search index",
            "icon": "🔍",
            "path": base_root / "db" / "3gpp_spec_search.db",
        },
        {
            "name": "3GPP Specification Change History",
            "subtitle": "CR/p-CR ledger & TDoc index",
            "icon": "📜",
            "path": base_root / "db" / "3gpp_spec_history.db",
            "can_wipe": True,
        },
    ]

    seen_paths = {c["path"].resolve() for c in candidate_files}
    for search_dir in [base_root / "db", base_root]:
        if search_dir.exists():
            for extra_db in search_dir.glob("*.db"):
                if extra_db.resolve() not in seen_paths and not extra_db.name.endswith("-wal"):
                    seen_paths.add(extra_db.resolve())
                    candidate_files.append({
                        "name": f"Auxiliary Database ({extra_db.name})",
                        "subtitle": "Custom or auxiliary storage file",
                        "icon": "🗄️",
                        "path": extra_db,
                    })

    return candidate_files


def query_db_metrics(db_path: Path) -> Dict[str, Any]:
    """Queries database size, freelist reclaimable space, table count, and WAL status."""
    if not db_path.exists():
        return {
            "exists": False,
            "size": 0,
            "wal_size": 0,
            "tables": 0,
            "reclaimable": 0,
            "freelist_pages": 0,
        }

    size = db_path.stat().st_size
    wal_path = Path(str(db_path) + "-wal")
    wal_size = wal_path.stat().st_size if wal_path.exists() else 0

    tables = 0
    reclaimable = 0
    freelist_pages = 0

    try:
        uri = db_path.resolve().as_uri() + "?mode=ro"
        conn = sqlite3.connect(uri, uri=True, timeout=1.0)
        try:
            fl_row = conn.execute("PRAGMA freelist_count;").fetchone()
            freelist_pages = fl_row[0] if fl_row else 0
            ps_row = conn.execute("PRAGMA page_size;").fetchone()
            page_size = ps_row[0] if ps_row else 4096
            reclaimable = freelist_pages * page_size

            tb_row = conn.execute(
                "SELECT count(*) FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
            ).fetchone()
            tables = tb_row[0] if tb_row else 0
        finally:
            conn.close()
    except Exception as e:
        logger.debug("Failed to query SQLite metrics for %s: %s", db_path.name, e)

    return {
        "exists": True,
        "size": size,
        "wal_size": wal_size,
        "tables": tables,
        "reclaimable": reclaimable,
        "freelist_pages": freelist_pages,
    }


def vacuum_database_file(db_path: Path) -> Tuple[bool, int, int, str]:
    """
    Executes WAL checkpointing and VACUUM on the specified SQLite database file.
    Returns: (success, bytes_before, bytes_after, message)
    """
    if not db_path.exists():
        return False, 0, 0, "File does not exist"

    wal_file = Path(str(db_path) + "-wal")
    bytes_before = db_path.stat().st_size + (wal_file.stat().st_size if wal_file.exists() else 0)

    try:
        conn = sqlite3.connect(str(db_path), check_same_thread=False)
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            conn.execute("VACUUM;")
            conn.execute("PRAGMA optimize;")
        finally:
            conn.close()

        bytes_after = db_path.stat().st_size + (wal_file.stat().st_size if wal_file.exists() else 0)
        return True, bytes_before, bytes_after, "Success"
    except Exception as e:
        return False, bytes_before, bytes_before, str(e)
