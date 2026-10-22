"""Read-oriented access to the 3GPP specification catalogue."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.specifications.core.database import SpecsDatabase


class SpecificationKnowledgeService:
    """Semantic read API over the existing Specifications database."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self._db: Optional[SpecsDatabase] = None

    def _database(self) -> SpecsDatabase:
        if self._db is None:
            if not self.db_path.exists():
                raise FileNotFoundError(str(self.db_path))
            self._db = SpecsDatabase(self.db_path)
        return self._db

    def find_specifications(self, query: str, limit: int = 8) -> ToolResult:
        query = str(query or "").strip()
        if not query:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Specification query must not be empty.")

        limit = _bounded_limit(limit)
        try:
            rows = self._db_search(query)
        except FileNotFoundError:
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE,
                message=f"Specifications database is unavailable: {self.db_path}",
            )
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Specification catalogue query failed: {exc}")

        if not rows:
            return ToolResult(
                ToolStatus.NOT_FOUND,
                data=[],
                message=f"No specification matched '{query}'.",
                metadata={"query": query, "returned": 0, "total_matches": 0, "more_available": False},
            )

        grouped: Dict[str, Dict[str, Any]] = {}
        for row in rows:
            normalized = _normalize_file_row(row)
            if not normalized:
                continue
            spec = normalized["specification"]
            entry = grouped.setdefault(
                spec,
                {
                    "specification": spec,
                    "type": normalized.get("type", "TS"),
                    "title": normalized.get("title", ""),
                    "versions": [],
                },
            )
            version = normalized.get("version")
            if version and version not in entry["versions"]:
                entry["versions"].append(version)

        matches = list(grouped.values())
        for item in matches:
            item["versions"] = _sort_versions(item["versions"], reverse=True)
            item["latest_known_version"] = item["versions"][0] if item["versions"] else None

        matches.sort(key=lambda item: _catalogue_rank(query, item))
        total = len(matches)
        returned = matches[:limit]

        return ToolResult(
            ToolStatus.FOUND,
            data=returned,
            metadata={
                "query": query,
                "returned": len(returned),
                "total_matches": total,
                "more_available": total > len(returned),
            },
        )

    def get_specification(self, specification: str) -> ToolResult:
        spec = _normalize_specification(specification)
        if not spec:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Specification number must not be empty.")

        try:
            db = self._database()
            details = db.get_spec_details(spec)
            file_rows = db.search_files(spec_number=spec)
        except FileNotFoundError:
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE,
                message=f"Specifications database is unavailable: {self.db_path}",
            )
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Specification lookup failed: {exc}")

        if not details and not file_rows:
            return ToolResult(
                ToolStatus.NOT_FOUND,
                message=f"Specification {spec} is not present in the catalogue.",
                metadata={"specification": spec},
            )

        normalized_files = [r for r in (_normalize_file_row(row) for row in file_rows) if r]
        versions = _sort_versions(
            [r["version"] for r in normalized_files if r.get("version")],
            reverse=True,
        )

        payload: Dict[str, Any] = {"specification": spec}
        if isinstance(details, dict):
            payload.update(details)
        payload["known_versions"] = versions
        payload["latest_known_version"] = versions[0] if versions else None
        payload["catalogue_file_count"] = len(normalized_files)

        return ToolResult(ToolStatus.FOUND, data=payload)

    def _db_search(self, query: str) -> Sequence[Any]:
        db = self._database()
        spec_candidate = _normalize_specification(query)
        return db.search_files(
            spec_number=spec_candidate if _looks_like_specification(spec_candidate) else query.strip()
        )


def _normalize_file_row(row: Any) -> Optional[Dict[str, Any]]:
    """Normalize the tuple shape currently returned by SpecsDatabase.search_files()."""
    if isinstance(row, dict):
        spec = row.get("spec_number") or row.get("number")
        if not spec:
            return None
        return {
            "specification": str(spec),
            "type": row.get("type") or "TS",
            "title": row.get("title") or "",
            "filename": row.get("filename") or "",
            "version": row.get("version") or "",
            "url": row.get("url") or row.get("file_url") or "",
            "upload_date": row.get("upload_date") or row.get("specification_upload_date") or "",
        }

    if isinstance(row, (tuple, list)) and len(row) >= 8:
        # SpecsDatabase.search_files() returns:
        # (s.name, sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date, tags, comments)
        tags_raw = row[8] if len(row) > 8 and row[8] else ""
        comments_raw = row[9] if len(row) > 9 and row[9] else ""
        item = {
            "specification": str(row[1]),
            "title": row[2] or "",
            "type": row[3] or "TS",
            "filename": row[4] or "",
            "version": row[5] or "",
            "url": row[6] or "",
            "upload_date": row[7] or "",
        }
        if tags_raw:
            item["tags"] = [t.strip() for t in tags_raw.split(",") if t.strip()]
        if comments_raw:
            item["user_comments"] = comments_raw
        if len(row) > 10 and row[10]:
            item["primary_group"] = row[10]
        return item
    return None


def _catalogue_rank(query: str, item: Dict[str, Any]) -> Tuple[int, str]:
    q = query.lower()
    spec = str(item.get("specification", "")).lower()
    title = str(item.get("title", "")).lower()
    tags = [t.lower() for t in item.get("tags", [])]
    comments = str(item.get("user_comments", "")).lower()

    if q == spec or q in (f"ts {spec}", f"tr {spec}"):
        score = 0
    elif any(q in t for t in tags):
        score = 1
    elif q in title:
        score = 2
    elif q in spec or q in comments:
        score = 3
    else:
        score = 4
    return score, spec


def _bounded_limit(limit: int) -> int:
    try:
        value = int(limit)
    except (TypeError, ValueError):
        value = 8
    return max(1, min(value, 20))


def _normalize_specification(value: str) -> str:
    text = str(value or "").strip().upper()
    return re.sub(r"^(?:3GPP\s*)?(?:(?:TS|TR)(?=\s|\d|$)\s*)?", "", text, flags=re.IGNORECASE).strip()


def _looks_like_specification(value: str) -> bool:
    return bool(re.match(r"^\d{2}\.\d{3}(?:-\d{1,2})?$", str(value or "").strip()))


def _version_tuple(version: str) -> Tuple[int, ...]:
    parts = []
    for token in str(version or "").lstrip("vV").split("."):
        if token.isdigit():
            parts.append(int(token))
        else:
            digits = "".join(ch for ch in token if ch.isdigit())
            parts.append(int(digits) if digits else -1)
    return tuple(parts)


def _sort_versions(versions: List[str], reverse: bool = False) -> List[str]:
    return sorted(set(str(v) for v in versions if v), key=_version_tuple, reverse=reverse)
