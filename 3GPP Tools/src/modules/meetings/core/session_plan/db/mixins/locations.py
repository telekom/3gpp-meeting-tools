# --- File: src/modules/meetings/core/session_plan/db/mixins/locations.py ---
"""Physical room & per-slot location notes with bidirectional meeting_locations.json sync."""

from datetime import datetime, timezone
import json
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple

logger = logging.getLogger("modules.meetings.core.session_plan.session_plan_db")


class SessionPlanLocationsMixin:
    """Manages physical room notes, per-slot overrides, and meeting_locations.json synchronization."""

    MEETING_LOCATIONS_FILENAME = "meeting_locations.json"

    def resolve_locations_json_path(self, meeting_dir: Optional[Path] = None) -> Optional[Path]:
        """Resolves `<meeting_root>/meeting_locations.json` alongside `meeting_timezone.json`."""
        base = Path(meeting_dir) if meeting_dir else self.db_path.parent
        mtg_root = base.parent if base.name.lower() == "agenda" else base
        return mtg_root / self.MEETING_LOCATIONS_FILENAME

    @classmethod
    def load_folder_locations_file(cls, meeting_dir: Optional[Path]) -> Tuple[Dict[str, str], Dict[str, str]]:
        """Reads `(room_locations, slot_locations)` directly from `<meeting_root>/meeting_locations.json`."""
        if not meeting_dir:
            return {}, {}
        base = Path(meeting_dir)
        mtg_root = base.parent if base.name.lower() == "agenda" else base
        loc_file = mtg_root / cls.MEETING_LOCATIONS_FILENAME
        if not loc_file.is_file():
            return {}, {}
        try:
            raw = json.loads(loc_file.read_text(encoding="utf-8"))
            r_map = {
                str(k).strip(): str(v).strip()
                for k, v in (raw.get("room_locations") or {}).items()
                if str(k).strip() and str(v).strip()
            }
            s_map = {
                str(k).strip(): str(v).strip()
                for k, v in (raw.get("slot_locations") or {}).items()
                if str(k).strip() and str(v).strip()
            }
            return r_map, s_map
        except Exception as e:
            logger.debug(f"[SessionPlanDB] Failed reading {loc_file}: {e}")
            return {}, {}

    def _sync_locations_from_json_if_needed(self, meeting_number: str, meeting_dir: Optional[Path] = None):
        """Seeds or updates SQLite location tables from `meeting_locations.json` when empty or externally updated."""
        loc_file = self.resolve_locations_json_path(meeting_dir)
        if not loc_file or not loc_file.is_file():
            return
        try:
            json_mtime = loc_file.stat().st_mtime
            db_mtime = self.db_path.stat().st_mtime if self.db_path.exists() else 0.0
            with self._get_connection() as conn:
                r_count = conn.execute(
                    "SELECT COUNT(*) FROM room_locations WHERE meeting_number = ?", (str(meeting_number),)
                ).fetchone()[0]
                s_count = conn.execute(
                    "SELECT COUNT(*) FROM slot_locations WHERE meeting_number = ?", (str(meeting_number),)
                ).fetchone()[0]
                has_db_data = (r_count + s_count) > 0
                if has_db_data and json_mtime <= db_mtime:
                    return

                r_map, s_map = self.load_folder_locations_file(loc_file.parent)
                if not r_map and not s_map and has_db_data and json_mtime <= db_mtime:
                    return
                conn.execute("DELETE FROM room_locations WHERE meeting_number = ?", (str(meeting_number),))
                conn.execute("DELETE FROM slot_locations WHERE meeting_number = ?", (str(meeting_number),))
                for rn, note in r_map.items():
                    conn.execute(
                        "INSERT INTO room_locations (meeting_number, room_name, location_note) VALUES (?, ?, ?)",
                        (str(meeting_number), rn, note),
                    )
                for sk, note in s_map.items():
                    conn.execute(
                        "INSERT INTO slot_locations (meeting_number, slot_key, location_note) VALUES (?, ?, ?)",
                        (str(meeting_number), sk, note),
                    )
        except Exception as e:
            logger.debug(f"[SessionPlanDB] Location JSON sync skipped: {e}")

    def _write_locations_json_conn(self, conn, meeting_number: str, meeting_dir: Optional[Path] = None):
        """Writes `<meeting_root>/meeting_locations.json` from the active transaction."""
        loc_file = self.resolve_locations_json_path(meeting_dir)
        if not loc_file:
            return
        try:
            r_rows = conn.execute(
                "SELECT room_name, location_note FROM room_locations WHERE meeting_number = ? ORDER BY room_name",
                (str(meeting_number),),
            ).fetchall()
            s_rows = conn.execute(
                "SELECT slot_key, location_note FROM slot_locations WHERE meeting_number = ? ORDER BY slot_key",
                (str(meeting_number),),
            ).fetchall()
            payload = {
                "meeting_number": str(meeting_number),
                "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "room_locations": {row["room_name"]: row["location_note"] for row in r_rows},
                "slot_locations": {row["slot_key"]: row["location_note"] for row in s_rows},
            }
            loc_file.parent.mkdir(parents=True, exist_ok=True)
            loc_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            logger.warning(f"[SessionPlanDB] Failed writing {loc_file}: {e}")

    def get_room_locations(self, meeting_number: str, meeting_dir: Optional[Path] = None) -> Dict[str, str]:
        self._sync_locations_from_json_if_needed(meeting_number, meeting_dir)
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT room_name, location_note FROM room_locations WHERE meeting_number = ? ORDER BY room_name",
                (str(meeting_number),),
            ).fetchall()
            return {row["room_name"]: row["location_note"] for row in rows}

    def get_slot_locations(self, meeting_number: str, meeting_dir: Optional[Path] = None) -> Dict[str, str]:
        self._sync_locations_from_json_if_needed(meeting_number, meeting_dir)
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT slot_key, location_note FROM slot_locations WHERE meeting_number = ? ORDER BY slot_key",
                (str(meeting_number),),
            ).fetchall()
            return {row["slot_key"]: row["location_note"] for row in rows}

    def set_room_location(
        self, meeting_number: str, room_name: str, location_note: str, meeting_dir: Optional[Path] = None
    ):
        clean_room = (room_name or "").strip()
        if not clean_room:
            return
        clean_note = (location_note or "").strip()
        with self._get_connection() as conn:
            if not clean_note:
                conn.execute(
                    "DELETE FROM room_locations WHERE meeting_number = ? AND room_name = ?",
                    (str(meeting_number), clean_room),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO room_locations (meeting_number, room_name, location_note, updated_at)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(meeting_number, room_name) DO UPDATE SET
                        location_note = excluded.location_note,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    (str(meeting_number), clean_room, clean_note),
                )
            self._write_locations_json_conn(conn, meeting_number, meeting_dir)

    def set_slot_location(
        self, meeting_number: str, slot_key: str, location_note: str, meeting_dir: Optional[Path] = None
    ):
        clean_key = (slot_key or "").strip()
        if not clean_key:
            return
        clean_note = (location_note or "").strip()
        with self._get_connection() as conn:
            if not clean_note:
                conn.execute(
                    "DELETE FROM slot_locations WHERE meeting_number = ? AND slot_key = ?",
                    (str(meeting_number), clean_key),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO slot_locations (meeting_number, slot_key, location_note, updated_at)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(meeting_number, slot_key) DO UPDATE SET
                        location_note = excluded.location_note,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    (str(meeting_number), clean_key, clean_note),
                )
            self._write_locations_json_conn(conn, meeting_number, meeting_dir)

    def set_all_room_locations(
        self,
        meeting_number: str,
        room_map: Dict[str, str],
        slot_map: Optional[Dict[str, str]] = None,
        meeting_dir: Optional[Path] = None,
    ):
        with self._get_connection() as conn:
            conn.execute("DELETE FROM room_locations WHERE meeting_number = ?", (str(meeting_number),))
            for rn, note in (room_map or {}).items():
                clean_rn = str(rn).strip()
                clean_note = str(note).strip()
                if clean_rn and clean_note:
                    conn.execute(
                        """
                        INSERT INTO room_locations (meeting_number, room_name, location_note, updated_at)
                        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                        """,
                        (str(meeting_number), clean_rn, clean_note),
                    )
            if slot_map is not None:
                conn.execute("DELETE FROM slot_locations WHERE meeting_number = ?", (str(meeting_number),))
                for sk, note in slot_map.items():
                    clean_sk = str(sk).strip()
                    clean_note = str(note).strip()
                    if clean_sk and clean_note:
                        conn.execute(
                            """
                            INSERT INTO slot_locations (meeting_number, slot_key, location_note, updated_at)
                            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                            """,
                            (str(meeting_number), clean_sk, clean_note),
                        )
            self._write_locations_json_conn(conn, meeting_number, meeting_dir)
