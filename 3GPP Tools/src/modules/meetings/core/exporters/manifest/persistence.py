# --- File: src/modules/meetings/core/exporters/manifest/persistence.py ---
"""Atomic persistence and legacy directory reconstruction for .export_manifest.json."""

import datetime
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.paths import to_extended_path

from .models import ExportManifest
from .validation import MANIFEST_FILENAME, _json_fallback, is_export_complete

logger = logging.getLogger(__name__)


def get_manifest_path(export_dir: Path) -> Path:
    """Returns standard path to the manifest file inside export_dir."""
    return export_dir / MANIFEST_FILENAME


def load_manifest(export_dir: Path) -> Optional[ExportManifest]:
    """Loads and deserializes .export_manifest.json if present."""
    m_path = get_manifest_path(export_dir)
    ext_path = to_extended_path(m_path)
    if not os.path.exists(ext_path):
        return None

    try:
        with open(ext_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return ExportManifest.from_dict(data)
    except Exception as e:
        logger.warning(f"[ExportManifest] Failed to load manifest from {m_path}: {e}")
        return None


def save_manifest(export_dir: Path, manifest: ExportManifest) -> bool:
    """Atomically serializes manifest to .export_manifest.json."""
    m_path = get_manifest_path(export_dir)
    ext_path = to_extended_path(m_path)
    temp_path = to_extended_path(export_dir / f"{MANIFEST_FILENAME}.tmp")

    try:
        os.makedirs(to_extended_path(export_dir), exist_ok=True)
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(manifest.to_dict(), f, indent=2, ensure_ascii=False, default=_json_fallback)
        os.replace(temp_path, ext_path)
        logger.info(f"[ExportManifest] Successfully saved manifest ({len(manifest.tdocs)} TDocs) to {m_path}")
        return True
    except Exception as e:
        logger.error(f"[ExportManifest] Failed to save manifest to {m_path}: {e}")
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
        return False


def reconstruct_manifest_from_directory(
    export_dir: Path,
    meeting_info: dict,
    all_rows: list,
    export_type: str = "pre_meeting",
) -> Optional[ExportManifest]:
    """
    Reconstructs an ExportManifest by inspecting existing files in an export directory
    that was created before .export_manifest.json existed.
    """
    if not export_dir.is_dir():
        return None

    master_index = export_dir / "index.md"
    if not master_index.is_file():
        return None

    wg = str(meeting_info.get("wg_name", "3GPP")).strip()
    mtg = str(meeting_info.get("meeting_number", "")).strip()
    mtg_title = f"{wg} #{mtg}" if (wg and mtg) else "3GPP Meeting"

    # Row lookup map by uppercase TDoc ID
    tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in all_rows if r.get("TDoc")}

    # Find creation timestamp from index.md modified time or header
    created_at = datetime.datetime.fromtimestamp(os.path.getmtime(to_extended_path(master_index))).strftime(
        "%Y-%m-%d %H:%M:%S"
    )
    try:
        with open(to_extended_path(master_index), "r", encoding="utf-8") as f:
            head = f.read(2048)
            m = re.search(r"\*\*Generated:\*\*\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})", head)
            if m:
                created_at = f"{m.group(1)}:00"
    except Exception:
        pass

    tdocs: Dict[str, Dict[str, Any]] = {}
    added_tdoc_ids: List[str] = []

    # Recursively find all *.md files under export_dir except index.md
    for root, _, files in os.walk(to_extended_path(export_dir)):
        for f in files:
            if not f.lower().endswith(".md") or f.lower() == "index.md":
                continue

            md_path = Path(root) / f
            # Skip empty or incomplete export files so they aren't registered as valid completed exports
            if not is_export_complete(md_path):
                continue
            try:
                rel_path = md_path.relative_to(export_dir).as_posix()
            except Exception:
                rel_path = md_path.name

            # Filename is sanitized TDoc ID
            tdoc_stem = md_path.stem
            tid_upper = tdoc_stem.upper()

            row = tdoc_map.get(tid_upper, {})
            parts = rel_path.split("/")
            rel_name = parts[0] if len(parts) >= 2 else "General"
            wi_name = parts[1] if len(parts) >= 3 else "General"
            cat_name = parts[2] if len(parts) >= 4 else "DISC"

            title = str(row.get("Title", "")).strip()
            source = str(row.get("Source", "")).strip()
            doc_type = str(row.get("Type", "")).strip()
            status = str(row.get("TDoc Status", "")).strip()
            ai_num = str(row.get("Agenda Item", "")).strip()

            # If row was not found in meeting rows, attempt reading frontmatter from markdown
            if not title and os.path.exists(to_extended_path(md_path)):
                try:
                    with open(to_extended_path(md_path), "r", encoding="utf-8") as mf:
                        content_head = mf.read(1024)
                        tm = re.search(r"^#\s+[A-Za-z0-9\-]+:\s*(.+)$", content_head, re.MULTILINE)
                        if tm:
                            title = tm.group(1).strip()
                        sm = re.search(r"\*\*Source:\*\*\s*([^\n\r]+)", content_head)
                        if sm:
                            source = sm.group(1).strip()
                except Exception:
                    pass

            tdocs[tid_upper] = {
                "tdoc_id": tid_upper,
                "title": title or tid_upper,
                "source": source,
                "type": doc_type or cat_name,
                "status": status,
                "ai": ai_num,
                "spec": str(row.get("Spec", "")).strip(),
                "release": rel_name,
                "wi": wi_name,
                "category": cat_name,
                "rel_path": rel_path,
                "session_id": 1,
                "added_at": created_at,
                "is_reserved": not is_export_complete(md_path),
                "my_status": str(row.get("My Status", "")),
                "my_notes": str(row.get("My Notes", "")),
                "supersedes": [],
                "replies": [],
                "reply_to_ids": [],
                "sibling_replies": [],
                "row_data": row,
            }
            added_tdoc_ids.append(tid_upper)

    if not tdocs:
        return None

    manifest = ExportManifest(
        meeting=mtg_title,
        export_type=export_type,
        created_at=created_at,
        last_updated_at=created_at,
        sessions=[
            {
                "session_id": 1,
                "timestamp": created_at,
                "label": "Initial Export (Reconstructed)",
                "scope_mode": "previous",
                "added_tdocs": sorted(added_tdoc_ids),
                "updated_tdocs": [],
                "total_tdocs": len(tdocs),
            }
        ],
        tdocs=tdocs,
    )
    # Save the reconstructed manifest so future runs find it immediately
    save_manifest(export_dir, manifest)
    return manifest


def load_or_reconstruct_manifest(
    export_dir: Path,
    meeting_info: dict,
    all_rows: list,
    export_type: str = "pre_meeting",
) -> Optional[ExportManifest]:
    """
    Attempts to load an existing manifest; if missing but directory contains an export,
    reconstructs it from disk. Returns None if export_dir does not contain a previous export.
    """
    manifest = load_manifest(export_dir)
    if manifest is not None:
        return manifest
    return reconstruct_manifest_from_directory(export_dir, meeting_info, all_rows, export_type=export_type)
