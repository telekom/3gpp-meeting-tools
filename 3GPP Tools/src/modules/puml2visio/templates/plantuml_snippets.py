# --- File: src/modules/puml2visio/templates/plantuml_snippets.py ---
"""Declarative catalog of PlantUML syntax elements, snippets, and 3GPP presets.

Supports dynamic placeholder resolution (participants, selection wrapping, cursor placement),
JSON configuration persistence (`src/config/plantuml_presets.json`), and git-ignored
user-defined presets (`src/config/plantuml_presets_user.json`).
"""

from dataclasses import dataclass, field
import json
import logging
from pathlib import Path
import re
from typing import Callable, Dict, List, Optional, Tuple

import core.utils.paths


from modules.puml2visio.templates.snippet_model import PlantUMLSnippet

# Regex to detect declared sequence participants/actors/entities/etc.
_PARTICIPANT_DECL_RE = re.compile(
    r'^\s*(?:participant|actor|boundary|control|entity|database|collections|queue)\s+'
    r'(?:"[^"]*"\s+as\s+([A-Za-z0-9_]+)|([A-Za-z0-9_]+))',
    re.MULTILINE
)

# Regex to detect implicit participants from arrow statements (A -> B)
_ARROW_PARTICIPANTS_RE = re.compile(
    r'^\s*([A-Za-z0-9_]+)\s*[-<~=]+(?:\[#[^\]]+\])?[-~=>x*o\\\/]+\s*([A-Za-z0-9_]+)\s*:',
    re.MULTILINE
)


def extract_participants(puml_text: str) -> List[str]:
    """Scans PlantUML text and returns an ordered, deduplicated list of participant aliases."""
    if not puml_text:
        return []

    participants: List[str] = []
    seen = set()

    def add(alias: str):
        if not alias:
            return
        if alias.lower() in {"start", "stop", "end", "return", "note", "activate", "deactivate", "alt", "opt"}:
            return
        if alias not in seen:
            seen.add(alias)
            participants.append(alias)

    # 1. Scan explicit declarations
    for match in _PARTICIPANT_DECL_RE.finditer(puml_text):
        alias = match.group(1) or match.group(2)
        if alias:
            add(alias.strip())

    # 2. Scan arrow message exchanges for any additional participants
    for match in _ARROW_PARTICIPANTS_RE.finditer(puml_text):
        p1 = match.group(1).strip()
        p2 = match.group(2).strip()
        add(p1)
        add(p2)

    return participants


def resolve_snippet_text(
    puml_text: str,
    snippet: PlantUMLSnippet,
    selected_text: str = "",
    base_indent: str = ""
) -> Tuple[str, int, int]:
    """Resolves dynamic placeholders in snippet code.

    Placeholders supported:
      {p1}: First detected participant (defaults to 'UE')
      {p2}: Second detected participant (defaults to 'gNB' or 'NW')
      {selection}: Selected text if wrapping, or default placeholder content
      {cursor}: Position where the editor cursor should be placed

    Returns:
      (resolved_text, selection_start_offset, selection_end_offset)
    """
    participants = extract_participants(puml_text)
    p1 = participants[0] if len(participants) > 0 else "UE"
    p2 = participants[1] if len(participants) > 1 else ("gNB" if p1 == "UE" else "NW")

    raw_code = snippet.code

    # Handle {selection} placeholder
    if "{selection}" in raw_code:
        if selected_text and snippet.is_block_wrapper:
            lines = selected_text.splitlines()
            indented_lines = ["  " + line if line.strip() else line for line in lines]
            inner_content = "\n".join(indented_lines)
        else:
            inner_content = "  ' TODO: Add messages or logic here"
        raw_code = raw_code.replace("{selection}", inner_content)

    # Replace participant tokens
    resolved = raw_code.replace("{p1}", p1).replace("{p2}", p2)

    # Indent multi-line snippets to match base_indent
    if base_indent:
        lines = resolved.splitlines()
        if lines:
            indented = [lines[0]] + [base_indent + l if l.strip() else l for l in lines[1:]]
            resolved = "\n".join(indented)

    # Locate cursor offset
    cursor_marker = "{cursor}"
    if cursor_marker in resolved:
        start_idx = resolved.find(cursor_marker)
        resolved = resolved.replace(cursor_marker, "", 1)
        end_marker = "{/cursor}"
        if end_marker in resolved:
            end_idx = resolved.find(end_marker)
            resolved = resolved.replace(end_marker, "", 1)
            return resolved, start_idx, end_idx
        return resolved, start_idx, start_idx

    return resolved, len(resolved), len(resolved)


from modules.puml2visio.templates.builtin_presets import (
    BUILTIN_SNIPPETS,
    SAMPLE_3GPP_PROCEDURES,
)


# =========================================================================

def get_default_presets_path() -> Path:
    """Returns path to the base, shipped presets configuration file (tracked in Git)."""
    return core.utils.paths.get_project_root() / "config" / "plantuml_presets.json"


def get_user_presets_path() -> Path:
    """Returns path to local workspace user presets file (strictly untracked in Git via .gitignore)."""
    return core.utils.paths.get_project_root() / "config" / "plantuml_presets_user.json"


def get_app_data_presets_path() -> Path:
    """Returns path to machine-level user presets file in user home app data root."""
    return core.utils.paths.get_app_data_root() / "plantuml_presets_user.json"


# =========================================================================
# --- PERSISTENCE & DYNAMIC HOT-RELOAD LOGIC ---
# =========================================================================

_RELOAD_CALLBACKS: List[Callable[[], None]] = []


def register_snippets_reloaded_callback(callback: Callable[[], None]):
    """Register a callback invoked whenever presets are reloaded or modified."""
    if callback not in _RELOAD_CALLBACKS:
        _RELOAD_CALLBACKS.append(callback)


def unregister_snippets_reloaded_callback(callback: Callable[[], None]):
    """Unregister a reload callback."""
    if callback in _RELOAD_CALLBACKS:
        _RELOAD_CALLBACKS.remove(callback)


def save_snippets_to_file(path: Path, snippets: List[PlantUMLSnippet]) -> None:
    """Writes a list of PlantUMLSnippet objects to a formatted JSON file."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "version": 1,
            "snippets": [s.to_dict() for s in snippets],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logging.error(f"Failed saving PlantUML snippets to {path}: {e}")


def load_snippets_from_file(path: Path, is_user_defined: bool = False) -> List[PlantUMLSnippet]:
    """Reads PlantUMLSnippet objects from a JSON file."""
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            raw_list = data.get("snippets", [])
            return [PlantUMLSnippet.from_dict(d, is_user_defined=is_user_defined) for d in raw_list]
    except Exception as e:
        logging.error(f"Failed reading PlantUML snippets from {path}: {e}")
        return []


def load_user_snippets_from_file(path: Path) -> List[PlantUMLSnippet]:
    """Reads user-defined snippets from a JSON file, flagging them as user defined."""
    return load_snippets_from_file(path, is_user_defined=True)


def save_user_snippets_to_file(path: Path, snippets: List[PlantUMLSnippet]) -> None:
    """Saves user-defined snippets to a JSON file."""
    save_snippets_to_file(path, snippets)


def _load_and_merge_user_file(
    path: Path,
    snippets_by_id: Dict[str, PlantUMLSnippet],
    ordered_list: List[PlantUMLSnippet]
):
    """Loads snippets from path and merges or appends them into ordered_list."""
    if not path.exists():
        return
    user_snips = load_user_snippets_from_file(path)
    for snip in user_snips:
        if snip.id in snippets_by_id:
            # Override existing built-in or previous snippet with user version
            idx = ordered_list.index(snippets_by_id[snip.id])
            ordered_list[idx] = snip
            snippets_by_id[snip.id] = snip
        else:
            snippets_by_id[snip.id] = snip
            ordered_list.append(snip)


def reload_snippets() -> List[PlantUMLSnippet]:
    """Loads default and user presets from disk, updating PLANTUML_SNIPPETS in-place.

    Loading Order:
      1. Default presets (`src/config/plantuml_presets.json`). If missing, creates it
         with all built-in 3GPP and UML presets.
      2. Workspace user presets (`src/config/plantuml_presets_user.json`, Git-ignored).
      3. AppData user presets (`~/3GPP_Delegate_Helper/plantuml_presets_user.json`).
    """
    default_path = get_default_presets_path()
    ordered_list: List[PlantUMLSnippet] = []
    snippets_by_id: Dict[str, PlantUMLSnippet] = {}

    if default_path.exists():
        loaded = load_snippets_from_file(default_path, is_user_defined=False)
        for s in loaded:
            snippets_by_id[s.id] = s
            ordered_list.append(s)

    # If default file is missing or empty, initialize with BUILTIN_SNIPPETS
    if not ordered_list:
        ordered_list = list(BUILTIN_SNIPPETS)
        for s in ordered_list:
            snippets_by_id[s.id] = s
        save_snippets_to_file(default_path, ordered_list)
    else:
        # Guarantee any newly added built-in presets (e.g. part_3gpp_ue) are present
        missing_builtins = [b for b in BUILTIN_SNIPPETS if b.id not in snippets_by_id]
        if missing_builtins:
            for b in missing_builtins:
                snippets_by_id[b.id] = b
                ordered_list.append(b)
            save_snippets_to_file(default_path, ordered_list)

    # Merge Workspace User Presets (plantuml_presets_user.json)
    user_path = get_user_presets_path()
    _load_and_merge_user_file(user_path, snippets_by_id, ordered_list)

    # Merge AppData User Presets
    app_data_path = get_app_data_presets_path()
    if app_data_path != user_path:
        _load_and_merge_user_file(app_data_path, snippets_by_id, ordered_list)

    # Update in-place so all imported references stay current
    PLANTUML_SNIPPETS[:] = ordered_list

    # Notify all registered GUI listeners to refresh their trees / tables / menus
    for cb in list(_RELOAD_CALLBACKS):
        try:
            cb()
        except RuntimeError as e:
            if "deleted" in str(e).lower() and cb in _RELOAD_CALLBACKS:
                _RELOAD_CALLBACKS.remove(cb)
            else:
                logging.error(f"Error in snippets reload callback: {e}")
        except Exception as e:
            logging.error(f"Error in snippets reload callback: {e}")

    return PLANTUML_SNIPPETS


def add_user_snippet(snippet: PlantUMLSnippet) -> None:
    """Adds or updates a user snippet in plantuml_presets_user.json and reloads."""
    user_path = get_user_presets_path()
    user_snippets = load_user_snippets_from_file(user_path)

    # Guarantee is_user_defined is True
    if not snippet.is_user_defined:
        snippet = PlantUMLSnippet(
            id=snippet.id,
            label=snippet.label,
            category=snippet.category,
            diagram_scope=snippet.diagram_scope,
            description=snippet.description,
            code=snippet.code,
            is_block_wrapper=snippet.is_block_wrapper,
            is_user_defined=True,
        )

    updated = False
    for i, s in enumerate(user_snippets):
        if s.id == snippet.id:
            user_snippets[i] = snippet
            updated = True
            break
    if not updated:
        user_snippets.append(snippet)

    save_user_snippets_to_file(user_path, user_snippets)
    reload_snippets()


def delete_user_snippet(snippet_id: str) -> bool:
    """Removes a user snippet from plantuml_presets_user.json and reloads."""
    user_path = get_user_presets_path()
    user_snippets = load_user_snippets_from_file(user_path)
    count_before = len(user_snippets)
    user_snippets = [s for s in user_snippets if s.id != snippet_id]
    if len(user_snippets) != count_before:
        save_user_snippets_to_file(user_path, user_snippets)
        reload_snippets()
        return True
    return False


# =========================================================================
# --- MODULE INITIALIZATION ---
# =========================================================================

# Global list of snippets imported by external modules:
# `from modules.puml2visio.templates.plantuml_snippets import PLANTUML_SNIPPETS`
PLANTUML_SNIPPETS: List[PlantUMLSnippet] = []

# Populate immediately on module load
reload_snippets()
