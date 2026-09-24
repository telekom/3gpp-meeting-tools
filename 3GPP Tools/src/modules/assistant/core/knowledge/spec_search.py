"""Progressive retrieval over the indexed specification-text database."""

from __future__ import annotations

import re
import threading
import uuid
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.spec_search.core.spec_search_db import SpecSearchDatabase, parse_version_tuple


class SpecSearchKnowledgeService:
    """Semantic, bounded read API over SpecSearchDatabase."""

    def __init__(self, db_path: Path, default_limit: int = 6, max_clause_chars: int = 24000):
        self.db_path = Path(db_path)
        self.default_limit = max(1, min(int(default_limit), 20))
        self.max_clause_chars = max(2000, int(max_clause_chars))
        self._db: Optional[SpecSearchDatabase] = None
        self._handles: Dict[str, int] = {}
        self._handle_lock = threading.RLock()

    def _database(self) -> SpecSearchDatabase:
        if self._db is None:
            if not self.db_path.exists():
                raise FileNotFoundError(str(self.db_path))
            self._db = SpecSearchDatabase(self.db_path)
        return self._db

    def get_indexed_versions(self, specification: str) -> ToolResult:
        spec = _normalize_specification(specification)
        if not spec:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Specification number must not be empty.")
        try:
            versions = self._database().get_versions_for_spec(spec)
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Specification search DB is unavailable: {self.db_path}")
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not inspect indexed versions: {exc}")

        if not versions:
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE,
                data=[],
                message=f"TS {spec} is not indexed in the specification search database.",
                metadata={"specification": spec},
            )
        return ToolResult(
            ToolStatus.FOUND,
            data=versions,
            metadata={"specification": spec, "latest_indexed_version": versions[0].get("version")},
        )

    def search_text(
        self,
        query: str,
        *,
        specification: Optional[str] = None,
        version: Optional[str] = None,
        release: Optional[int] = None,
        clause: Optional[str] = None,
        limit: Optional[int] = None,
    ) -> ToolResult:
        query = str(query or "").strip()
        if not query:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Search query must not be empty.")

        spec = _normalize_specification(specification) if specification else None
        try:
            db = self._database()
            selected_versions = self._resolve_versions(db, spec, version, release)
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Specification search DB is unavailable: {self.db_path}")
        except ValueError as exc:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=str(exc))
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not resolve specification search scope: {exc}")

        if not selected_versions:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message="No indexed specification versions match the requested scope.")

        version_ids = [int(v["id"]) for v in selected_versions if v.get("id") is not None]
        if not version_ids:
            return ToolResult(ToolStatus.ERROR, message="Indexed version records did not contain usable version IDs.")

        try:
            raw = db.search_substring(query, version_ids, clause_filter=clause or None)
            records = _records(raw)
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Specification text search failed: {exc}")

        if spec:
            records = [r for r in records if str(r.get("spec_number", "")) == spec]

        if not records:
            return ToolResult(
                ToolStatus.NOT_FOUND,
                data=[],
                message=f"No indexed clauses matched '{query}'.",
                metadata={
                    "query": query,
                    "specification": spec,
                    "searched_versions": [v.get("version") for v in selected_versions],
                    "returned": 0,
                    "total_matches": 0,
                    "more_available": False,
                },
            )

        records.sort(key=lambda r: _search_rank(query, r))
        max_results = _bounded_limit(limit if limit is not None else self.default_limit)
        selected = records[:max_results]

        output = []
        for record in selected:
            clause_pk = record.get("clause_pk")
            if clause_pk is None:
                continue
            result_id = self._register_handle(int(clause_pk))
            output.append(
                {
                    "result_id": result_id,
                    "specification": str(record.get("spec_number", "")),
                    "version": str(record.get("version", "")),
                    "release_date": record.get("release_date") or "",
                    "clause": str(record.get("clause_number", "")),
                    "title": record.get("clause_title") or "",
                    "snippet": record.get("snippet_text") or "",
                }
            )

        return ToolResult(
            ToolStatus.FOUND,
            data=output,
            metadata={
                "query": query,
                "specification": spec,
                "searched_versions": [v.get("version") for v in selected_versions],
                "returned": len(output),
                "total_matches": len(records),
                "more_available": len(records) > len(output),
            },
        )

    def get_clause(self, result_id: str) -> ToolResult:
        handle = str(result_id or "").strip()
        if not handle:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Search result ID must not be empty.")

        with self._handle_lock:
            clause_pk = self._handles.get(handle)
        if clause_pk is None:
            return ToolResult(
                ToolStatus.INVALID_REQUEST,
                message=f"Unknown or expired specification search result ID: {handle}",
            )

        try:
            data = self._database().get_clause_content(clause_pk)
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Specification search DB is unavailable: {self.db_path}")
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Clause retrieval failed: {exc}")

        if not data:
            return ToolResult(ToolStatus.NOT_FOUND, message=f"The clause referenced by {handle} no longer exists.")

        payload = dict(data)
        content = str(payload.get("content", "") or "")
        complete = len(content) <= self.max_clause_chars
        if not complete:
            content = content[: self.max_clause_chars]
        payload["content"] = content
        payload["content_complete"] = complete
        payload["result_id"] = handle

        metadata = {
            "original_content_chars": len(str(data.get("content", "") or "")),
            "returned_content_chars": len(content),
            "more_available": not complete,
        }
        return ToolResult(
            ToolStatus.FOUND if complete else ToolStatus.SOURCE_INCOMPLETE,
            data=payload,
            message="" if complete else "Clause is larger than the configured evidence payload and was returned partially.",
            metadata=metadata,
        )

    def _resolve_versions(
        self,
        db: SpecSearchDatabase,
        specification: Optional[str],
        version: Optional[str],
        release: Optional[int],
    ) -> List[Dict[str, Any]]:
        if version and not specification:
            raise ValueError("An exact specification version requires a specification number.")

        if specification:
            available = list(db.get_versions_for_spec(specification) or [])
            if not available:
                raise ValueError(f"TS {specification} has no indexed text.")

            if version:
                target = str(version).lstrip("vV").strip()
                matches = [v for v in available if str(v.get("version", "")).lstrip("vV") == target]
                if not matches:
                    raise ValueError(f"TS {specification} v{target} is not indexed.")
                return matches[:1]

            if release is not None:
                rel = int(release)
                matches = [v for v in available if _major_version(v.get("version")) == rel]
                if not matches:
                    raise ValueError(f"TS {specification} has no indexed Release {rel} version.")
                matches.sort(key=lambda v: parse_version_tuple(str(v.get("version", ""))), reverse=True)
                return matches[:1]

            # get_versions_for_spec is already newest-first, but sort defensively.
            available.sort(key=lambda v: parse_version_tuple(str(v.get("version", ""))), reverse=True)
            return available[:1]

        # No specification was requested: search the latest indexed version of
        # every indexed specification, avoiding historical duplicate matches.
        imported = list(db.get_imported_versions() or [])
        by_spec: Dict[str, List[Dict[str, Any]]] = {}
        for item in imported:
            by_spec.setdefault(str(item.get("spec_number", "")), []).append(item)

        selected = []
        for spec, items in by_spec.items():
            if not spec:
                continue
            if release is not None:
                items = [i for i in items if _major_version(i.get("version")) == int(release)]
            if not items:
                continue
            items.sort(key=lambda v: parse_version_tuple(str(v.get("version", ""))), reverse=True)
            selected.append(items[0])
        return selected

    def _register_handle(self, clause_pk: int) -> str:
        handle = f"S-{uuid.uuid4().hex[:12]}"
        with self._handle_lock:
            self._handles[handle] = clause_pk
        return handle


def _records(value: Any) -> List[Dict[str, Any]]:
    if value is None:
        return []
    if hasattr(value, "to_dict"):
        try:
            return list(value.to_dict(orient="records"))
        except TypeError:
            pass
    if isinstance(value, list):
        return [dict(v) for v in value if isinstance(v, dict)]
    return []


def _search_rank(query: str, record: Dict[str, Any]) -> Tuple[int, int, int, str]:
    q = query.casefold()
    title = str(record.get("clause_title", "") or "").casefold()
    snippet = _strip_marks(str(record.get("snippet_text", "") or "")).casefold()
    tokens = [t for t in re.findall(r"[a-z0-9]+", q) if len(t) > 1]

    if q and q == title.strip():
        title_score = 0
    elif q and q in title:
        title_score = 1
    elif tokens and all(t in title for t in tokens):
        title_score = 2
    elif q and q in snippet:
        title_score = 3
    else:
        title_score = 4

    occurrence_score = -(snippet.count(q) if q else 0)
    order = int(record.get("order_index") or 0)
    clause = str(record.get("clause_number", "") or "")
    return title_score, occurrence_score, order, clause


def _strip_marks(text: str) -> str:
    return text.replace("<mark>", "").replace("</mark>", "")


def _bounded_limit(limit: int) -> int:
    try:
        value = int(limit)
    except (TypeError, ValueError):
        value = 6
    return max(1, min(value, 20))


def _normalize_specification(value: Optional[str]) -> str:
    text = str(value or "").strip().upper()
    for prefix in ("3GPP", "TS", "TR"):
        text = text.replace(prefix, "")
    return text.strip()


def _major_version(version: Any) -> Optional[int]:
    parsed = parse_version_tuple(str(version or ""))
    return parsed[0] if parsed else None
