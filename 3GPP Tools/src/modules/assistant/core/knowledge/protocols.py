"""Structured protocol knowledge service over the existing protocol database."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.assistant.core.protocol_registry import (
    ProtocolDescriptor,
    iter_protocols,
    resolve_protocol,
)
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple


class ProtocolKnowledgeService:
    """Semantic read API over the application's unified protocol database."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self._db: Optional[NASDatabase] = None

    def _database(self) -> NASDatabase:
        if self._db is None:
            if not self.db_path.exists():
                raise FileNotFoundError(str(self.db_path))
            self._db = NASDatabase(self.db_path)
        return self._db

    def list_protocols(self) -> ToolResult:
        try:
            imported = self._database().get_imported_versions()
        except FileNotFoundError:
            imported = []
            db_available = False
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not inspect protocol coverage: {exc}")
        else:
            db_available = True

        coverage = _group_imported(imported)
        data = []
        for descriptor in iter_protocols():
            imported_versions: List[str] = []
            for spec in descriptor.specifications:
                imported_versions.extend(v["version"] for v in coverage.get(spec, []))
            imported_versions = _sort_versions(imported_versions, reverse=True)
            data.append(
                {
                    "id": descriptor.id,
                    "protocol": descriptor.display_name,
                    "specifications": list(descriptor.specifications),
                    "parser_supported": descriptor.parser_supported,
                    "structured_data_available": bool(imported_versions),
                    "imported_versions": imported_versions,
                    "latest_imported_version": imported_versions[0] if imported_versions else None,
                    "reference_points": list(descriptor.reference_points),
                    "notes": descriptor.notes,
                }
            )

        status = ToolStatus.FOUND if db_available else ToolStatus.SOURCE_UNAVAILABLE
        message = "" if db_available else f"Protocol DB is unavailable: {self.db_path}"
        return ToolResult(status, data=data, message=message)

    def list_messages(self, protocol: str, *, version=None, release=None, limit: int = 100) -> ToolResult:
        """List messages/PDUs for a protocol without requiring a known message name."""
        protocol = str(protocol or "").strip()
        if not protocol:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Protocol name must not be empty.")
        descriptors = self._resolve_descriptors(protocol)
        if not descriptors:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=f"Unknown protocol: {protocol}")
        try:
            db = self._database()
            imported = db.get_imported_versions()
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Protocol DB is unavailable: {self.db_path}")
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not inspect protocol database: {exc}")

        selected = self._select_versions(imported, descriptors, version, release)
        if not selected:
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE, data=[],
                message=f"No structured protocol data is imported for {protocol}.",
                metadata={"protocol": protocol, "parser_supported": all(d.parser_supported for d in descriptors)},
            )
        version_ids = [int(v["id"]) for v in selected if v.get("id") is not None]
        try:
            rows = db.get_messages_list(version_ids)
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Protocol message listing failed: {exc}")

        items, seen = [], set()
        for row in rows or []:
            if not isinstance(row, dict):
                continue
            name = str(row.get("message_name", "") or "").strip()
            if not name:
                continue
            key = (name.casefold(), str(row.get("spec_number","")), str(row.get("version","")))
            if key in seen:
                continue
            seen.add(key)
            items.append({
                "message": name,
                "clause": str(row.get("clause", "") or ""),
                "specification": str(row.get("spec_number", "") or ""),
                "version": str(row.get("version", "") or ""),
            })
        items.sort(key=lambda x: x["message"].casefold())
        try:
            max_results = max(1, min(int(limit), 250))
        except (TypeError, ValueError):
            max_results = 100
        total = len(items)
        returned = items[:max_results]
        if not returned:
            return ToolResult(ToolStatus.NOT_FOUND, data=[],
                              message=f"No structured messages are stored for {protocol}.",
                              metadata={"searched_versions": _version_labels(selected)})
        return ToolResult(
            ToolStatus.FOUND,
            data={
                "protocol": descriptors[0].display_name if len(descriptors) == 1 else protocol,
                "specifications": sorted({x["specification"] for x in returned if x["specification"]}),
                "versions": _sort_versions([x["version"] for x in returned if x["version"]], reverse=True),
                "messages": returned,
                "source_kind": "structured_protocol_knowledge",
            },
            metadata={"searched_versions": _version_labels(selected), "returned": len(returned),
                      "total_messages": total, "more_available": total > len(returned)},
        )

    def find_message(
        self,
        message: str,
        *,
        protocol: Optional[str] = None,
        version: Optional[str] = None,
        release: Optional[int] = None,
    ) -> ToolResult:
        message = str(message or "").strip()
        if not message:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Protocol message name must not be empty.")

        try:
            db = self._database()
            imported = db.get_imported_versions()
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Protocol DB is unavailable: {self.db_path}")
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not inspect protocol database: {exc}")

        descriptors = self._resolve_descriptors(protocol)
        if protocol and not descriptors:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=f"Unknown protocol: {protocol}")

        selected = self._select_versions(imported, descriptors, version, release)
        if not selected:
            label = protocol or "requested protocol scope"
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE,
                message=f"No structured protocol data is imported for {label}.",
                metadata={"protocol": protocol, "parser_supported": bool(descriptors and all(d.parser_supported for d in descriptors))},
            )

        version_ids = [int(v["id"]) for v in selected if v.get("id") is not None]
        try:
            messages = db.get_messages_list(version_ids)
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Protocol message lookup failed: {exc}")

        candidates = [m for m in messages if _message_matches(message, m.get("message_name", ""))]
        if not candidates:
            return ToolResult(
                ToolStatus.NOT_FOUND,
                data=[],
                message=f"No structured protocol message matched '{message}'.",
                metadata={"searched_versions": _version_labels(selected)},
            )

        ranked = sorted(candidates, key=lambda m: _name_rank(message, str(m.get("message_name", ""))))
        best_rank = _name_rank(message, str(ranked[0].get("message_name", "")))[0]
        best = [m for m in ranked if _name_rank(message, str(m.get("message_name", "")))[0] == best_rank]

        if len(best) > 1 and best_rank > 0:
            return ToolResult(
                ToolStatus.AMBIGUOUS,
                data=best[:10],
                message=f"Multiple protocol messages match '{message}'.",
                metadata={"total_matches": len(best), "searched_versions": _version_labels(selected)},
            )

        target = best[0]
        target_name = str(target.get("message_name", ""))
        try:
            df = db.get_message_evolution_df(
                message_name=target_name,
                version_ids=version_ids,
                include_descriptions=True,
            )
            fields = _records(df)
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not retrieve fields for '{target_name}': {exc}")

        spec_numbers = sorted({str(r.get("spec_number")) for r in fields if r.get("spec_number")})
        versions = _sort_versions([str(r.get("version")) for r in fields if r.get("version")], reverse=True)
        clause = str(target.get("clause", "") or "")
        payload = {
            "message": target_name,
            "clause": clause,
            "specifications": spec_numbers,
            "versions": versions,
            "fields": fields,
            "source_kind": "structured_protocol_knowledge",
        }
        return ToolResult(
            ToolStatus.FOUND,
            data=payload,
            metadata={"searched_versions": _version_labels(selected), "field_count": len(fields)},
        )

    def find_ie(
        self,
        ie: str,
        *,
        protocol: Optional[str] = None,
        version: Optional[str] = None,
        release: Optional[int] = None,
        search_descriptions: bool = True,
    ) -> ToolResult:
        ie = str(ie or "").strip()
        if not ie:
            return ToolResult(ToolStatus.INVALID_REQUEST, message="Information Element query must not be empty.")

        try:
            db = self._database()
            imported = db.get_imported_versions()
        except FileNotFoundError:
            return ToolResult(ToolStatus.SOURCE_UNAVAILABLE, message=f"Protocol DB is unavailable: {self.db_path}")
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Could not inspect protocol database: {exc}")

        descriptors = self._resolve_descriptors(protocol)
        if protocol and not descriptors:
            return ToolResult(ToolStatus.INVALID_REQUEST, message=f"Unknown protocol: {protocol}")

        selected = self._select_versions(imported, descriptors, version, release)
        if not selected:
            return ToolResult(
                ToolStatus.SOURCE_UNAVAILABLE,
                message=f"No structured protocol data is imported for {protocol or 'the requested scope'}.",
            )

        version_ids = [int(v["id"]) for v in selected if v.get("id") is not None]
        try:
            messages = db.get_messages_by_ie_search(
                ie_query=ie,
                version_ids=version_ids,
                search_descriptions=bool(search_descriptions),
            )
        except Exception as exc:
            return ToolResult(ToolStatus.ERROR, message=f"Protocol IE search failed: {exc}")

        if not messages:
            return ToolResult(
                ToolStatus.NOT_FOUND,
                data=[],
                message=f"No structured protocol IE/field matched '{ie}'.",
                metadata={"searched_versions": _version_labels(selected)},
            )

        # Resolve definitions opportunistically from the matching message fields.
        definitions: List[Dict[str, Any]] = []
        seen_defs = set()
        for message_row in messages[:20]:
            message_name = str(message_row.get("message_name", "") or "")
            if not message_name:
                continue
            try:
                df = db.get_message_evolution_df(
                    message_name=message_name,
                    version_ids=version_ids,
                    include_descriptions=True,
                )
            except Exception:
                continue
            for field in _records(df):
                haystacks = (
                    str(field.get("ie_name", "")),
                    str(field.get("field_path", "")),
                    str(field.get("type_reference", "")),
                )
                if not any(_normalized(ie) in _normalized(value) for value in haystacks):
                    continue
                clause = str(field.get("type_reference", "") or field.get("ie_name", ""))
                spec_number = str(field.get("spec_number", "") or "")
                try:
                    defs = db.get_ie_definitions_by_clause(
                        clause=clause,
                        alt_name=str(field.get("ie_name", "") or ie),
                        spec_number=spec_number or None,
                        version_ids=version_ids,
                    )
                except Exception:
                    defs = []
                for definition in defs:
                    key = (
                        definition.get("spec_number"),
                        definition.get("version"),
                        definition.get("clause"),
                        definition.get("ie_name"),
                    )
                    if key not in seen_defs:
                        seen_defs.add(key)
                        definitions.append(definition)

        return ToolResult(
            ToolStatus.FOUND,
            data={
                "query": ie,
                "matching_messages": messages,
                "definitions": definitions,
                "source_kind": "structured_protocol_knowledge",
            },
            metadata={
                "searched_versions": _version_labels(selected),
                "matching_message_count": len(messages),
                "definition_count": len(definitions),
            },
        )

    @staticmethod
    def _resolve_descriptors(protocol: Optional[str]) -> Tuple[ProtocolDescriptor, ...]:
        if not protocol:
            return tuple(iter_protocols())
        return resolve_protocol(protocol)

    @staticmethod
    def _select_versions(
        imported: Sequence[Dict[str, Any]],
        descriptors: Tuple[ProtocolDescriptor, ...],
        version: Optional[str],
        release: Optional[int],
    ) -> List[Dict[str, Any]]:
        allowed_specs = {s for d in descriptors for s in d.specifications}
        rows = [dict(v) for v in imported if str(v.get("spec_number", "")) in allowed_specs]

        if version:
            target = str(version).lstrip("vV").strip()
            rows = [v for v in rows if str(v.get("version", "")).lstrip("vV") == target]
        if release is not None:
            rows = [v for v in rows if _major(v.get("version")) == int(release)]

        by_spec: Dict[str, List[Dict[str, Any]]] = {}
        for row in rows:
            by_spec.setdefault(str(row.get("spec_number", "")), []).append(row)

        selected = []
        for spec_rows in by_spec.values():
            spec_rows.sort(key=lambda v: parse_version_tuple(str(v.get("version", ""))), reverse=True)
            selected.append(spec_rows[0])
        return selected


def _group_imported(rows: Sequence[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    result: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        result.setdefault(str(row.get("spec_number", "")), []).append(dict(row))
    return result


def _message_matches(query: str, candidate: str) -> bool:
    q = _normalized(query)
    c = _normalized(candidate)
    return bool(q and c and (q == c or q in c or c in q))


def _name_rank(query: str, candidate: str) -> Tuple[int, int, str]:
    q = _normalized(query)
    c = _normalized(candidate)
    if q == c:
        score = 0
    elif c.startswith(q) or q.startswith(c):
        score = 1
    elif q in c:
        score = 2
    else:
        score = 3
    return score, abs(len(c) - len(q)), candidate.casefold()


def _normalized(value: str) -> str:
    return "".join(ch.casefold() for ch in str(value or "") if ch.isalnum())


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


def _major(version: Any) -> Optional[int]:
    parsed = parse_version_tuple(str(version or ""))
    return parsed[0] if parsed else None


def _sort_versions(versions: List[str], reverse: bool = False) -> List[str]:
    return sorted(set(versions), key=lambda v: parse_version_tuple(str(v)), reverse=reverse)


def _version_labels(rows: Sequence[Dict[str, Any]]) -> List[str]:
    return [f"TS {r.get('spec_number')} v{r.get('version')}" for r in rows]
