# --- File: src/modules/meetings/core/exporters/export_manifest.py ---
"""
Persistent Export Manifest & Changelog Tracker for 3GPP Meeting Briefing Dossiers.

Maintains state in '.export_manifest.json' within the export directory, enabling:
1. Fast incremental / append exports (skipping downloads and conversions for already exported TDocs).
2. Non-destructive corpus merging (re-running with narrower or wider filters preserves prior work).
3. Human-readable and LLM-optimized delta changelogs at the bottom of index.md files.
"""

import datetime
import json
import logging
import os
import posixpath
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.tdoc_markdown_converter import sanitize_filename

logger = logging.getLogger(__name__)

MANIFEST_FILENAME = ".export_manifest.json"
MANIFEST_VERSION = 1


def _clean_for_json(obj: Any) -> Any:
    """Recursively converts sets, tuples, and Paths into JSON-serializable types."""
    if isinstance(obj, (set, frozenset)):
        return sorted([_clean_for_json(x) for x in obj], key=lambda x: str(x))
    if isinstance(obj, (list, tuple)):
        return [_clean_for_json(x) for x in obj]
    if isinstance(obj, dict):
        return {str(k): _clean_for_json(v) for k, v in obj.items()}
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()
    return obj


def _json_fallback(obj: Any) -> Any:
    """Fallback handler for json.dump default parameter."""
    if isinstance(obj, (set, frozenset)):
        return sorted(list(obj), key=lambda x: str(x))
    if isinstance(obj, tuple):
        return list(obj)
    if isinstance(obj, Path):
        return str(obj)
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    if hasattr(obj, "__dict__"):
        return obj.__dict__
    return str(obj)


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem and URL-breaking characters."""
    clean = re.sub(r'[\\/*?:"<>|#%]', "_", str(name or "")).strip()
    return re.sub(r'_+', '_', clean).strip("._ ") or "General"


def is_export_complete(
    md_path: Union[str, Path],
    doc_info: Optional[Dict[str, Any]] = None,
    source_path: Optional[Union[str, Path]] = None,
) -> bool:
    """
    Validates whether an existing Markdown export file on disk is complete and usable.
    Returns False if:
      - File does not exist or is 0 bytes
      - File size is suspiciously small (< 120 bytes, e.g. truncated header or empty stub)
      - File is an empty placeholder stub ("Document conversion was not available")
      - File is a presentation export (PPTX/PPT) but has no slides ("## Slide ")
      - File is a presentation export referencing slide images that are missing on disk
    """
    p = Path(to_extended_path(md_path))
    if not p.is_file():
        return False
    try:
        st = p.stat()
        if st.st_size == 0:
            return False

        with open(p, "r", encoding="utf-8", errors="replace") as f:
            head = f.read(2048)

        # Minimum content threshold: reject tiny stubs unless stamped by converter
        if st.st_size < 50:
            return False
        if st.st_size < 120 and "<!--- tdoc_converter:" not in head:
            return False

        # Placeholder stub created when source was missing
        if "Document conversion was not available" in head:
            return False

        # Placeholder stub created when document was reserved / not uploaded
        if (
            "Status: Reserved / Not Uploaded" in head
            or "allocated or reserved on the meeting agenda but has not yet been uploaded" in head
            or "⏳ **Status: Reserved" in head
        ):
            return False

        # Corrupted image prefix from previous bug
        if "__-________" in head:
            return False

        # Detect presentation export
        is_pres = False
        if source_path and Path(source_path).suffix.lower() in (".pptx", ".ppt"):
            is_pres = True
        elif doc_info:
            src_file = str(doc_info.get("original_file", "") or doc_info.get("file", "") or doc_info.get("filename", "")).lower()
            if src_file.endswith((".pptx", ".ppt")):
                is_pres = True
            doc_type = str(doc_info.get("type", "")).lower()
            if "presentation" in doc_type or "slides" in doc_type:
                is_pres = True

        if not is_pres:
            if "type=presentation" in head or "mode=presentation" in head:
                is_pres = True
            elif re.search(r"\*\*Original file:\*\*.*?\.ppt[x]?", head, re.IGNORECASE):
                is_pres = True
            elif "- **Slides:**" in head:
                is_pres = True

        # Read content to validate completeness and image presence
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(65536)

        if "__-________" in content:
            return False

        if is_pres and "## Slide " not in content:
            return False

        # Check all referenced images exist on disk
        all_imgs = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", content)
        for img_rel in all_imgs:
            img_rel = img_rel.strip().replace("\\", "/")
            if img_rel.startswith(("http://", "https://", "data:")):
                continue
            img_path = (p.parent / img_rel).resolve()
            if not img_path.is_file():
                # Check fallback locations: {stem}_media/ and media/
                img_name = posixpath.basename(img_rel)
                stem_media = (p.parent / f"{p.stem}_media" / img_name).resolve()
                sub_media = (p.parent / "media" / img_name).resolve()
                if not stem_media.is_file() and not sub_media.is_file():
                    return False

        return True
    except Exception:
        return False


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

    def to_dict(self) -> Dict[str, Any]:
        return {
            "manifest_version": self.manifest_version,
            "export_type": self.export_type,
            "meeting": self.meeting,
            "created_at": self.created_at,
            "last_updated_at": self.last_updated_at,
            "sessions": _clean_for_json(self.sessions),
            "tdocs": _clean_for_json(self.tdocs),
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
            manifest_version=data.get("manifest_version", MANIFEST_VERSION),
        )

    def is_tdoc_exported(self, tdoc_id: str) -> bool:
        return str(tdoc_id).strip().upper() in self.tdocs

    def get_total_tdocs(self) -> int:
        return len(self.tdocs)

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
        session_record = {
            "session_id": session_id,
            "timestamp": timestamp,
            "label": label,
            "scope_mode": scope_mode,
            "added_tdocs": sorted(list(added_tdocs)),
            "updated_tdocs": sorted(list(updated_tdocs)),
            "total_tdocs": len(self.tdocs) + len(added_tdocs),
        }
        self.sessions.append(session_record)
        self.last_updated_at = timestamp

        for entry in added_entries:
            tid = str(entry.get("tdoc_id", "")).strip().upper()
            if not tid:
                continue
            self.tdocs[tid] = {
                "tdoc_id": tid,
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
    created_at = datetime.datetime.fromtimestamp(os.path.getmtime(to_extended_path(master_index))).strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(to_extended_path(master_index), "r", encoding="utf-8") as f:
            head = f.read(2048)
            m = re.search(r'\*\*Generated:\*\*\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})', head)
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
                        tm = re.search(r'^#\s+[A-Za-z0-9\-]+:\s*(.+)$', content_head, re.MULTILINE)
                        if tm:
                            title = tm.group(1).strip()
                        sm = re.search(r'\*\*Source:\*\*\s*([^\n\r]+)', content_head)
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


def plan_incremental_session(
    manifest: Optional[ExportManifest],
    incoming_entries: dict,
    export_dir: Path,
    force_reconvert: bool = False,
) -> IncrementalPlan:
    """
    Analyzes incoming export entries against existing manifest and disk files to plan
    an incremental session.
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if manifest is None or not manifest.tdocs:
        # Fresh / Initial export
        new_tdocs = set()
        for rel in incoming_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        new_tdocs.add(str(doc["tdoc_id"]).strip().upper())
        return IncrementalPlan(
            is_incremental=False,
            session_id=1,
            session_timestamp=now_str,
            new_tdocs=new_tdocs,
            tdocs_to_skip=set(),
            tdocs_to_update=set(),
            preserved_tdocs=set(),
        )

    # Incremental Session
    session_id = len(manifest.sessions) + 1
    new_tdocs: Set[str] = set()
    tdocs_to_skip: Set[str] = set()
    tdocs_to_update: Set[str] = set()

    for rel_name, wis in incoming_entries.items():
        rel_slug = sanitize_folder_name(rel_name)
        for wi_name, wi_data in wis.items():
            wi_slug = sanitize_folder_name(wi_name)
            for cat_name, cat_docs in wi_data.get("categories", {}).items():
                cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                type_dir = export_dir / rel_slug / wi_slug / cat_folder
                for doc in cat_docs:
                    tid = str(doc["tdoc_id"]).strip().upper()
                    target_abs = doc.get("target_md_abs_path")
                    if not target_abs:
                        target_abs = type_dir / f"{sanitize_filename(tid)}.md"

                    # Check if target file already exists on disk
                    file_exists = False
                    existing_path = None
                    if target_abs and os.path.exists(to_extended_path(target_abs)):
                        file_exists = True
                        existing_path = Path(to_extended_path(target_abs))
                    elif tid in manifest.tdocs:
                        saved_rel = manifest.tdocs[tid].get("rel_path")
                        if saved_rel and os.path.exists(to_extended_path(export_dir / saved_rel)):
                            file_exists = True
                            existing_path = Path(to_extended_path(export_dir / saved_rel))

                    # Completeness check: is the existing export non-empty, complete, and valid?
                    complete = is_export_complete(existing_path, doc) if (file_exists and existing_path) else False

                    if tid in tdocs_to_update:
                        continue
                    if file_exists and complete and not force_reconvert:
                        tdocs_to_skip.add(tid)
                    elif file_exists and (not complete or force_reconvert):
                        tdocs_to_skip.discard(tid)
                        tdocs_to_update.add(tid)
                    else:
                        tdocs_to_skip.discard(tid)
                        new_tdocs.add(tid)

    # All TDocs already in manifest that were not in incoming entries are preserved
    incoming_all_tids = {
        str(d["tdoc_id"]).strip().upper()
        for rel in incoming_entries.values()
        for wi in rel.values()
        for cat_docs in wi["categories"].values()
        for d in cat_docs
    }
    preserved_tdocs = set(manifest.tdocs.keys()) - incoming_all_tids

    return IncrementalPlan(
        is_incremental=True,
        session_id=session_id,
        session_timestamp=now_str,
        new_tdocs=new_tdocs,
        tdocs_to_skip=tdocs_to_skip,
        tdocs_to_update=tdocs_to_update,
        preserved_tdocs=preserved_tdocs,
    )


def merge_manifest_into_tree(
    manifest: ExportManifest,
    tree: dict,
    export_dir: Path,
) -> dict:
    """
    Merges preserved documents from manifest into the export tree so master and local
    index.md reports display the complete corpus without accidental loss.
    """
    if not manifest or not manifest.tdocs:
        return tree

    # Set of all TDoc IDs already in tree
    existing_in_tree = {
        str(d["tdoc_id"]).strip().upper()
        for rel in tree.values()
        for wi in rel.values()
        for cat_docs in wi["categories"].values()
        for d in cat_docs
    }

    for tid, doc_info in manifest.tdocs.items():
        if tid in existing_in_tree:
            continue

        rel_name = doc_info.get("release") or "General"
        wi_name = doc_info.get("wi") or "General"
        cat_name = doc_info.get("category") or "DISC"
        rel_path = doc_info.get("rel_path") or f"{rel_name}/{wi_name}/{cat_name}/{sanitize_filename(tid)}.md"
        target_abs = export_dir / rel_path
        if not target_abs.is_file() or not is_export_complete(target_abs):
            continue

        if rel_name not in tree:
            tree[rel_name] = {}
        if wi_name not in tree[rel_name]:
            tree[rel_name][wi_name] = {
                "info": {"name": wi_name},
                "ais": set(),
                "ai_titles": {},
                "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
            }

        ai_num = doc_info.get("ai") or "Unassigned"
        if ai_num and ai_num != "Unassigned":
            tree[rel_name][wi_name]["ais"].add(ai_num)

        entry = {
            "tdoc_id": tid,
            "title": doc_info.get("title", ""),
            "source": doc_info.get("source", ""),
            "type": doc_info.get("type", ""),
            "status": doc_info.get("status", ""),
            "ai": ai_num,
            "spec": doc_info.get("spec", ""),
            "category": cat_name,
            "cross_wis": [],
            "replies": doc_info.get("replies", []),
            "reply_to_ids": doc_info.get("reply_to_ids", []),
            "sibling_replies": doc_info.get("sibling_replies", []),
            "supersedes": doc_info.get("supersedes", []),
            "is_reserved": not is_export_complete(target_abs, doc_info) if os.path.exists(to_extended_path(target_abs)) else bool(doc_info.get("is_reserved", False)),
            "my_status": doc_info.get("my_status", ""),
            "my_notes": doc_info.get("my_notes", ""),
            "for_action": "",
            "sanitized_companies": [],
            "row_data": doc_info.get("row_data", {}),
            "tdoc_url": "",
            "target_md_abs_path": target_abs,
            "target_md_rel_path": rel_path,
            "wi_local_rel_path": f"{cat_name}/{sanitize_filename(tid)}.md",
            "is_preserved_from_previous": True,
        }

        if cat_name not in tree[rel_name][wi_name]["categories"]:
            tree[rel_name][wi_name]["categories"][cat_name] = []
        tree[rel_name][wi_name]["categories"][cat_name].append(entry)

    # Sort entries within categories
    for rel in tree.values():
        for wi in rel.values():
            for cat in wi["categories"]:
                wi["categories"][cat].sort(key=lambda x: str(x.get("tdoc_id", "")))

    return tree


def format_master_changelog_markdown(
    manifest: ExportManifest,
    current_session: Dict[str, Any],
    new_doc_entries: List[Dict[str, Any]],
    updated_doc_entries: Optional[List[Dict[str, Any]]] = None,
    base_dir: Optional[Path] = None,
) -> str:
    """
    Renders an unobtrusive HTML <details> collapsible changelog block at the bottom
    of the master index.md file, optimized for human reading and LLM incremental analysis.
    """
    updated_doc_entries = updated_doc_entries or []
    sessions = manifest.sessions or [current_session]
    total_docs = manifest.get_total_tdocs()

    is_update = current_session.get("session_id", 1) > 1
    added_count = len(new_doc_entries)
    updated_count = len(updated_doc_entries)
    timestamp = current_session.get("timestamp", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

    if is_update:
        delta_str = f"+{added_count} TDoc(s) added"
        if updated_count:
            delta_str += f", {updated_count} updated"
        summary_title = (
            f"🕒 <b>Export Changelog & LLM Incremental Update Guide</b> — "
            f"<i>Latest update: {timestamp} ({delta_str}, {total_docs} total)</i>"
        )
    else:
        summary_title = (
            f"🕒 <b>Export Changelog & LLM Incremental Update Guide</b> — "
            f"<i>Initial Export ({total_docs} TDocs)</i>"
        )

    lines = [
        "\n---\n",
        "## 🔄 Export History & Incremental LLM Guide\n",
        "<details>",
        f"<summary>{summary_title}</summary>\n",
        "### 📅 Export Sessions History\n",
    ]

    for s in sessions:
        sid = s.get("session_id", 1)
        st_time = s.get("timestamp", "")
        s_lbl = s.get("label", f"Session #{sid}")
        s_added = len(s.get("added_tdocs", []))
        s_tot = s.get("total_tdocs", 0)
        s_scope = s.get("scope_mode", "")
        scope_suffix = f" *(scope: `{s_scope}`)*" if s_scope else ""
        if sid == 1:
            lines.append(f"- **Session #1 ({s_lbl}):** `{st_time}` — {s_added} initial TDoc(s){scope_suffix}")
        else:
            lines.append(f"- **Session #{sid} ({s_lbl}):** `{st_time}` — **+{s_added} added** ({s_tot} total in dossier){scope_suffix}")

    lines.append("")

    # Delta table for newly added TDocs in this session
    if new_doc_entries:
        lines.append(f"### 🆕 Newly Added TDocs in Latest Update (+{len(new_doc_entries)})\n")
        lines.append("| TDoc | Type | Work Item | AI | Title | Source | Dossier Link |")
        lines.append("|:---|:---|:---|:---:|:---|:---|:---:|")

        for d in sorted(new_doc_entries, key=lambda x: str(x.get("tdoc_id", ""))):
            tid = d.get("tdoc_id", "")
            doc_type = d.get("type", "")
            wi = d.get("wi", "General")
            ai = d.get("ai", "")
            title = str(d.get("title", "")).replace("\n", " ").replace("|", "\\|")
            source = str(d.get("source", "")).replace("\n", " ").replace("|", "\\|")
            rel_link = d.get("rel_path") or d.get("target_md_rel_path") or ""

            link_cell = f"[📄 Open]({rel_link})" if rel_link else tid
            lines.append(f"| `{tid}` | {doc_type} | `{wi}` | {ai} | {title} | {source} | {link_cell} |")

        lines.append("")

    # Updated / Revised TDocs table (if any)
    if updated_doc_entries:
        lines.append(f"### 🔄 Updated / Re-converted TDocs ({len(updated_doc_entries)})\n")
        lines.append("| TDoc | Type | Work Item | Title | Source | Link |")
        lines.append("|:---|:---|:---|:---|:---|:---:|")
        for d in sorted(updated_doc_entries, key=lambda x: str(x.get("tdoc_id", ""))):
            tid = d.get("tdoc_id", "")
            doc_type = d.get("type", "")
            wi = d.get("wi", "General")
            title = str(d.get("title", "")).replace("\n", " ").replace("|", "\\|")
            source = str(d.get("source", "")).replace("\n", " ").replace("|", "\\|")
            rel_link = d.get("rel_path") or d.get("target_md_rel_path") or ""
            link_cell = f"[📄 Open]({rel_link})" if rel_link else tid
            lines.append(f"| `{tid}` | {doc_type} | `{wi}` | {title} | {source} | {link_cell} |")
        lines.append("")

    # Actionable LLM Instruction Guide Prompt
    lines.append("#### 💡 Suggested Instruction for LLM Analysis:\n")
    if is_update and new_doc_entries:
        new_ids_str = ", ".join(f"`{d['tdoc_id']}`" for d in new_doc_entries[:15])
        if len(new_doc_entries) > 15:
            new_ids_str += f", and {len(new_doc_entries) - 15} more"
        lines.append(
            "> *Copy/paste this prompt to your LLM to perform an incremental analysis without re-analyzing previous documents:*\n"
            ">\n"
            f"> \"I have extended this meeting briefing with **{len(new_doc_entries)} additional TDocs** ({new_ids_str}) "
            "as detailed in the Changelog table above. Without repeating or regenerating the analysis of previously summarized "
            "contributions, please perform an incremental analysis focusing strictly on:\n"
            "> 1. **New Proposals:** Summarize the core proposals and architectural objectives in the newly added TDocs.\n"
            "> 2. **Cross-Impact & Consensus Shifts:** Identify if these new contributions alter, resolve, or contradict "
            "discussions/stances from the earlier documents.\n"
            "> 3. **Action Items & Stances:** Note any specific operator/vendor consensus changes or recommendations for "
            "our delegation based on these additions.\""
        )
    else:
        lines.append(
            "> *Copy/paste this prompt when initiating analysis on this briefing dossier:*\n"
            ">\n"
            f"> \"Analyze this 3GPP meeting briefing dossier ({total_docs} contributions). "
            "Synthesize the key debate clusters per Work Item, identify competing proposals, highlight vendor alignments, "
            "and suggest priority items for delegate preparation.\""
        )

    lines.append("</details>\n")
    return "\n".join(lines)


def format_local_wi_changelog_markdown(
    wi_name: str,
    added_docs_in_wi: List[Dict[str, Any]],
    timestamp: str,
) -> str:
    """Renders a compact collapsible changelog block inside a local Work Item index.md."""
    if not added_docs_in_wi:
        return ""

    n = len(added_docs_in_wi)
    lines = [
        "\n---\n",
        "<details>",
        f"<summary>🕒 <b>Work Item Incremental Update (+{n} TDoc(s) added on {timestamp})</b></summary>\n",
        f"The following {n} contribution(s) were added to this Work Item briefing a posteriori:\n",
    ]

    for d in sorted(added_docs_in_wi, key=lambda x: str(x.get("tdoc_id", ""))):
        tid = d.get("tdoc_id", "")
        title = d.get("title", "")
        source = d.get("source", "")
        cat = d.get("category", "")
        src_str = f" ({source})" if source else ""
        lines.append(f"- **`{tid}`** [{cat}]: *{title}*{src_str}")

    lines.append("\n</details>\n")
    return "\n".join(lines)
