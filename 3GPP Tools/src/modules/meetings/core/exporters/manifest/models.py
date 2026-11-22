# --- File: src/modules/meetings/core/exporters/manifest/models.py ---
"""Data models for persistent export manifests and incremental export plans."""

import datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set

from .validation import MANIFEST_VERSION, _clean_for_json


@dataclass
class IncrementalPlan:
    """Plan for an export run detailing what is new, skipped, or updated."""

    is_incremental: bool
    session_id: int
    session_timestamp: str
    new_tdocs: Set[str] = field(default_factory=set)
    tdocs_to_skip: Set[str] = field(default_factory=set)
    tdocs_to_update: Set[str] = field(default_factory=set)
    preserved_tdocs: Set[str] = field(default_factory=set)


class ExportManifest:
    """Represents the persistent manifest of an exported dossier directory."""

    def __init__(
        self,
        meeting: str = "",
        export_type: str = "pre_meeting",
        created_at: Optional[str] = None,
        last_updated_at: Optional[str] = None,
        sessions: Optional[List[Dict[str, Any]]] = None,
        tdocs: Optional[Dict[str, Dict[str, Any]]] = None,
        exported_meetings: Optional[Dict[str, Dict[str, Any]]] = None,
        exported_decisions: Optional[Dict[str, str]] = None,
        synthesis_state: Optional[Dict[str, Any]] = None,
        manifest_version: int = MANIFEST_VERSION,
    ):
        self.manifest_version = manifest_version
        self.meeting = meeting
        self.export_type = export_type
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.created_at = created_at or now_str
        self.last_updated_at = last_updated_at or self.created_at
        self.sessions = sessions or []
        self.tdocs = tdocs or {}
        self.exported_meetings = exported_meetings or {}
        self.exported_decisions = exported_decisions or {}
        self.synthesis_state = synthesis_state or {}

    def to_dict(self) -> Dict[str, Any]:
        return {
            "manifest_version": self.manifest_version,
            "export_type": self.export_type,
            "meeting": self.meeting,
            "created_at": self.created_at,
            "last_updated_at": self.last_updated_at,
            "sessions": _clean_for_json(self.sessions),
            "tdocs": _clean_for_json(self.tdocs),
            "exported_meetings": _clean_for_json(self.exported_meetings),
            "exported_decisions": _clean_for_json(self.exported_decisions),
            "synthesis_state": _clean_for_json(self.synthesis_state),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ExportManifest":
        return cls(
            meeting=str(data.get("meeting", "")),
            export_type=str(data.get("export_type", "pre_meeting")),
            created_at=data.get("created_at"),
            last_updated_at=data.get("last_updated_at"),
            sessions=data.get("sessions", []),
            tdocs=data.get("tdocs", {}),
            exported_meetings=data.get("exported_meetings", {}),
            exported_decisions=data.get("exported_decisions", {}),
            synthesis_state=data.get("synthesis_state", {}),
            manifest_version=data.get("manifest_version", 1),
        )

    def is_tdoc_exported(self, tdoc_id: str) -> bool:
        return str(tdoc_id).strip().upper() in self.tdocs

    def get_total_tdocs(self) -> int:
        return len(self.tdocs)

    def is_decision_exported(self, decision_key: str) -> bool:
        return str(decision_key) in self.exported_decisions

    def record_decision(self, decision_key: str, rel_path: str) -> None:
        self.exported_decisions[str(decision_key)] = str(rel_path or "").replace("\\", "/")

    def register_tdoc_entries(
        self,
        entries: List[Dict[str, Any]],
        session_id: int,
        timestamp: str,
    ) -> None:
        """Upserts TDoc metadata entries into self.tdocs immediately (used by checkpoints and record_session)."""
        for entry in entries:
            tid = str(entry.get("tdoc_id", "")).strip().upper()
            if not tid:
                continue
            self.tdocs[tid] = {
                "tdoc_id": tid,
                "wg": str(entry.get("wg", "")).strip().upper(),
                "meeting_number": str(entry.get("meeting_number", "")).strip(),
                "meeting_folder_id": str(entry.get("meeting_folder_id", "")).strip(),
                "title": str(entry.get("title", "")).strip(),
                "source": str(entry.get("source", "")).strip(),
                "type": str(entry.get("type", "")).strip(),
                "status": str(entry.get("status", "")).strip(),
                "ai": str(entry.get("ai", "")).strip(),
                "spec": str(entry.get("spec", "")).strip(),
                "release": str(entry.get("release", "General")).strip(),
                "wi": str(entry.get("wi", "General")).strip(),
                "category": str(entry.get("category", "DISC")).strip(),
                "rel_path": str(entry.get("rel_path", "")).replace("\\", "/"),
                "session_id": session_id,
                "added_at": timestamp,
                "is_reserved": bool(entry.get("is_reserved", False)),
                "my_status": str(entry.get("my_status", "")),
                "my_notes": str(entry.get("my_notes", "")),
                "supersedes": list(entry.get("supersedes", [])),
                "replies": list(entry.get("replies", [])),
                "reply_to_ids": list(entry.get("reply_to_ids", [])),
                "sibling_replies": list(entry.get("sibling_replies", [])),
                "row_data": entry.get("row_data", {}),
            }

    def record_meeting_checkpoint(
        self,
        meeting_key: str,
        checkpoint_data: Dict[str, Any],
        entries: Optional[List[Dict[str, Any]]] = None,
        session_id: int = 1,
        timestamp: Optional[str] = None,
    ) -> None:
        """Records per-meeting completion status and registers processed TDocs immediately for resumability."""
        ts = timestamp or datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.last_updated_at = ts
        if entries:
            self.register_tdoc_entries(entries, session_id=session_id, timestamp=ts)
        payload = dict(checkpoint_data)
        payload.setdefault("completed_at", ts)
        self.exported_meetings[meeting_key] = payload

    def record_session(
        self,
        session_id: int,
        timestamp: str,
        scope_mode: str,
        added_tdocs: List[str],
        updated_tdocs: List[str],
        added_entries: List[Dict[str, Any]],
    ) -> None:
        """Records a new export session and updates registered TDoc metadata."""
        label = "Initial Export" if session_id == 1 else f"Incremental Update #{session_id - 1}"
        self.last_updated_at = timestamp
        self.register_tdoc_entries(added_entries, session_id=session_id, timestamp=timestamp)
        session_record = {
            "session_id": session_id,
            "timestamp": timestamp,
            "label": label,
            "scope_mode": scope_mode,
            "added_tdocs": sorted(list(added_tdocs)),
            "updated_tdocs": sorted(list(updated_tdocs)),
            "total_tdocs": len(self.tdocs),
        }
        self.sessions.append(session_record)
