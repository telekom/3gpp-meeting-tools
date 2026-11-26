"""Shared argument helpers, Win32 COM locks, DB path resolution, and cache root lookup."""

from __future__ import annotations

import faulthandler
import os
import re
import shutil
import sys
import threading
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from core.utils.paths import get_cache_root, get_db_root, get_meetings_root
from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.assistant.core.tool_registry import ArgumentType, ToolArgument


def ToolArgumentSpec(
    name: str,
    arg_type: str,
    description: str,
    *,
    required: bool = False,
    default: Any = None,
    minimum: Optional[int] = None,
    maximum: Optional[int] = None,
    max_length: Optional[int] = None,
    enum: Tuple[Any, ...] = (),
) -> ToolArgument:
    """Helper constructing a `ToolArgument` from compact string type names."""
    type_map = {
        "string": ArgumentType.STRING,
        "integer": ArgumentType.INTEGER,
        "boolean": ArgumentType.BOOLEAN,
    }
    return ToolArgument(
        name=name,
        type=type_map.get(arg_type, ArgumentType.STRING),
        description=description,
        required=required,
        default=default,
        minimum=minimum,
        maximum=maximum,
        max_length=max_length,
        enum=tuple(enum),
    )


def _wrap_handler(fn: Callable[[Dict[str, Any]], Any]) -> Callable[..., ToolResult]:
    """Adapts a `fn(args_dict) -> Dict | ToolResult` handler to `fn(**kwargs) -> ToolResult`."""
    def _wrapped(**kwargs: Any) -> ToolResult:
        res = fn(dict(kwargs))
        if isinstance(res, ToolResult):
            is_ok = res.status in (ToolStatus.FOUND, ToolStatus.SOURCE_INCOMPLETE)
            if isinstance(res.data, dict):
                payload = dict(res.data)
                payload.setdefault("ok", is_ok)
                if not is_ok and res.message and "error" not in payload:
                    payload["error"] = res.message
                if res.metadata:
                    payload.setdefault("metadata", res.metadata)
            elif isinstance(res.data, list):
                payload = {
                    "ok": is_ok,
                    "items": res.data,
                    "count": len(res.data),
                }
                if not is_ok and res.message:
                    payload["error"] = res.message
                if res.metadata:
                    payload["metadata"] = res.metadata
            else:
                payload = {
                    "ok": is_ok,
                    "result": res.data,
                }
                if not is_ok and res.message:
                    payload["error"] = res.message
                if res.metadata:
                    payload["metadata"] = res.metadata
            return ToolResult(
                status=ToolStatus.FOUND,
                data=payload,
                message=res.message,
                evidence_ids=res.evidence_ids,
                metadata=res.metadata,
            )
        if isinstance(res, dict):
            if "error" in res and "ok" not in res:
                res["ok"] = False
            elif "ok" not in res:
                res["ok"] = True
            return ToolResult(
                status=ToolStatus.FOUND,
                data=res,
                message=str(res.get("error") or "ok"),
            )
        return ToolResult(status=ToolStatus.FOUND, data={"ok": True, "result": res})
    return _wrapped

# Global re-entrant lock for Win32 COM automation (Word, Visio, PowerPoint)
_COM_LOCK = threading.RLock()


def _disable_faulthandler_for_com() -> None:
    """Disables faulthandler on Windows so recoverable Win32 COM RPC_E_CALL_REJECTED SEH exceptions
    are caught cleanly by Python try/except blocks instead of printing alarming stderr tracebacks."""
    if os.name == "nt":
        try:
            faulthandler.disable()
        except Exception:
            pass


def _resolve_db_paths(
    specs_db_path: Optional[Path] = None,
    spec_search_db_path: Optional[Path] = None,
    protocol_db_path: Optional[Path] = None,
    spec_history_db_path: Optional[Path] = None,
    user_tdocs_db_path: Optional[Path] = None,
    session_plan_db_path: Optional[Path] = None,
) -> Dict[str, Path]:
    db_root = get_db_root()
    return {
        "specs_db": Path(specs_db_path or (db_root / "3gpp_data.db")),
        "spec_search_db": Path(spec_search_db_path or (db_root / "3gpp_spec_search.db")),
        "protocol_db": Path(protocol_db_path or (db_root / "3gpp_protocol_data.db")),
        "spec_history_db": Path(spec_history_db_path or (db_root / "3gpp_spec_history.db")),
        "user_tdocs_db": Path(user_tdocs_db_path or (db_root / "user_tdocs.db")),
        "session_plan_db": Path(session_plan_db_path or (db_root / "session_plans.db")),
    }


def _parse_csv_list(val: Any) -> List[str]:
    if not val:
        return []
    if isinstance(val, (list, tuple, set)):
        return [str(x).strip() for x in val if str(x).strip()]
    return [x.strip() for x in re.split(r"[,;]+", str(val)) if x.strip()]


def _extract_resolved_word_docs(resolved: Any) -> List[Path]:
    """Extracts a list of Word document Paths from `SpecDocResolver.resolve_spec_files()` output."""
    if isinstance(resolved, dict):
        raw = resolved.get("word_docs") or []
    elif isinstance(resolved, (list, tuple)):
        raw = resolved
    else:
        raw = []
    return [Path(p) for p in raw if p]


def _match_heading_for_clause(headings: Sequence[Any], clause_number: str) -> Optional[Any]:
    """Matches a `HeadingInfo` whose text begins with the exact clause number."""
    clean_num = str(clause_number or "").strip()
    if not clean_num:
        return None
    pattern = re.compile(r"^" + re.escape(clean_num) + r"(?:[\s\t]|$)", re.IGNORECASE)
    for h in headings:
        text = str(getattr(h, "text", "") or "").strip()
        if pattern.match(text):
            return h
    return None


def _export_doc_to_markdown(
    doc_path: Path,
    out_md: Optional[Path] = None,
    *,
    revisions: str = "markup",
    tdoc_id: Optional[str] = None,
) -> Tuple[Optional[Path], str, Any]:
    """Converts a `.docx` file to Markdown via `convert_document_to_markdown` using a directory `out_dir`."""
    from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

    target_dir = out_md.parent if out_md is not None else doc_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    target_id = out_md.stem if out_md is not None else (tdoc_id or doc_path.stem)
    exported = convert_document_to_markdown(
        doc_path,
        out_dir=target_dir,
        tdoc_id=target_id,
        revisions=revisions,
    )
    if not exported:
        return None, "", None
    produced = Path(exported.md_path) if getattr(exported, "md_path", None) else None
    if out_md is not None and produced and produced.exists() and produced.resolve() != out_md.resolve():
        shutil.move(str(produced), str(out_md))
        produced = out_md
    md_text = getattr(exported, "text", "") or ""
    if not md_text and produced and produced.exists():
        md_text = produced.read_text(encoding="utf-8", errors="replace")
    return produced, md_text, exported


def _get_meetings_cache_roots() -> List[Path]:
    """Returns candidate meeting cache roots in priority order."""
    roots: List[Path] = []
    try:
        from modules.meetings.core.settings import MeetingsSettings
        cfg_dir = Path(MeetingsSettings().cache_dir)
        if cfg_dir not in roots:
            roots.append(cfg_dir)
    except Exception:
        pass
    for fallback in (get_cache_root(), get_meetings_root()):
        if fallback not in roots:
            roots.append(fallback)
    return roots


_ORIG_GET_MEETINGS_CACHE_ROOTS = _get_meetings_cache_roots


def resolve_meetings_cache_roots() -> List[Path]:
    """Resolves meeting cache roots, honoring monkeypatches on the legacy external_agent_tools shim."""
    for mod_name in (
        "modules.assistant.core.tools.external_agent_tools",
        "src.modules.assistant.core.tools.external_agent_tools",
    ):
        shim = sys.modules.get(mod_name)
        if shim is not None:
            shim_fn = getattr(shim, "_get_meetings_cache_roots", None)
            if shim_fn is not None and shim_fn is not _ORIG_GET_MEETINGS_CACHE_ROOTS:
                return shim_fn()
    return _get_meetings_cache_roots()
