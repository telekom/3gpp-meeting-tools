# --- File: src/core/ui/hex_migrator.py ---
"""
AST/tokenizer-safe migration and hex-lint scanning utilities for 3GPP Delegate Tools.

Phase 0.4: Automates migrating common hardcoded `#RRGGBB` color literals in UI
modules to centralized `core.ui.theme.C` tokens, with `--dry-run` unified diff
preview before any edits are written.

Phase 0.5: Shared UI-file classification and raw-hex counting used by the
unittest raw-hex ratchet lint guard (`tests/test_ui_hex_lint.py`).
"""

from __future__ import annotations

import argparse
import ast
import difflib
import io
import re
import sys
import tokenize
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

from core.ui.theme import HEX_TO_TOKEN

HEX_COLOR_RE = re.compile(r"#[0-9A-Fa-f]{6}\b")

# Files under src/ that intentionally define tokens or run migration tooling.
LINT_EXEMPT_POSIX_PATHS: Set[str] = {
    "src/core/ui/theme.py",
    "src/core/ui/hex_migrator.py",
}


@dataclass
class MigrationResult:
    """Outcome of migrating a single Python source string or file."""

    path: Optional[Path]
    original: str
    updated: str
    replacements: int = 0
    unmapped_hexes: Dict[str, int] = field(default_factory=dict)
    skipped_tokens: int = 0
    import_added: bool = False

    @property
    def changed(self) -> bool:
        return self.updated != self.original


def _to_repo_posix(path: Path, repo_root: Path) -> str:
    """Return a normalized `src/...` POSIX path relative to `repo_root`."""
    try:
        return path.resolve().relative_to(repo_root.resolve()).as_posix()
    except ValueError:
        return path.as_posix()


def is_ui_python_file(repo_rel_posix: str, text: Optional[str] = None) -> bool:
    """
    Determine whether a repository-relative POSIX path is a UI Python module
    subject to design-token governance.
    """
    if not repo_rel_posix.endswith(".py"):
        return False
    if repo_rel_posix in LINT_EXEMPT_POSIX_PATHS:
        return False
    if (
        repo_rel_posix.startswith("src/core/ui/")
        or repo_rel_posix in ("src/main_window.py", "src/main_tools.py")
        or "/ui/" in repo_rel_posix
    ):
        return True
    if text is not None and ("PyQt5.QtWidgets" in text or "from PyQt5 import" in text):
        return True
    return False


def _collect_docstring_spans(tree: ast.AST) -> Set[Tuple[int, int, int, int]]:
    """Collect (lineno, col_offset, end_lineno, end_col_offset) spans of module/class/fn docstrings."""
    spans: Set[Tuple[int, int, int, int]] = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.body and isinstance(node.body[0], ast.Expr):
                val = node.body[0].value
                if isinstance(val, ast.Constant) and isinstance(val.value, str):
                    if None not in (val.lineno, val.col_offset, val.end_lineno, val.end_col_offset):
                        spans.add((val.lineno, val.col_offset, val.end_lineno, val.end_col_offset))
    return spans


def _span_in_docstrings(
    start: Tuple[int, int],
    end: Tuple[int, int],
    doc_spans: Set[Tuple[int, int, int, int]],
) -> bool:
    """Return True if a token span falls inside any collected docstring span."""
    for d_sl, d_sc, d_el, d_ec in doc_spans:
        if (start[0], start[1]) >= (d_sl, d_sc) and (end[0], end[1]) <= (d_el, d_ec):
            return True
    return False


def find_hex_literals_in_source(source: str) -> List[Tuple[int, str]]:
    """
    Return all `(line_number, '#RRGGBB')` occurrences inside runtime Python string
    tokens (excluding comments and module/class/function docstrings).
    """
    tree = ast.parse(source)
    doc_spans = _collect_docstring_spans(tree)
    findings: List[Tuple[int, str]] = []

    fstring_middle_type = getattr(tokenize, "FSTRING_MIDDLE", None)
    for tok in tokenize.generate_tokens(io.StringIO(source).readline):
        if tok.type == tokenize.STRING:
            if _span_in_docstrings(tok.start, tok.end, doc_spans):
                continue
            for m in HEX_COLOR_RE.finditer(tok.string):
                line_offset = tok.string[: m.start()].count("\n")
                findings.append((tok.start[0] + line_offset, m.group(0).upper()))
        elif fstring_middle_type is not None and tok.type == fstring_middle_type:
            if _span_in_docstrings(tok.start, tok.end, doc_spans):
                continue
            for m in HEX_COLOR_RE.finditer(tok.string):
                line_offset = tok.string[: m.start()].count("\n")
                findings.append((tok.start[0] + line_offset, m.group(0).upper()))
    return findings


def scan_ui_hex_counts(repo_root: Path) -> Dict[str, List[Tuple[int, str]]]:
    """Scan all UI Python files under `<repo_root>/src` and return `{posix_path: [(line, hex), ...]}`."""
    src_root = repo_root / "src"
    results: Dict[str, List[Tuple[int, str]]] = {}
    for path in sorted(src_root.rglob("*.py")):
        rel_posix = _to_repo_posix(path, repo_root)
        text = path.read_text(encoding="utf-8", errors="replace")
        if not is_ui_python_file(rel_posix, text=text):
            continue
        findings = find_hex_literals_in_source(text)
        if findings:
            results[rel_posix] = findings
    return results


def _split_string_prefix_and_quotes(token_str: str) -> Optional[Tuple[str, str, str]]:
    """Split a STRING token into `(prefix, quote, body)`."""
    idx = 0
    while idx < len(token_str) and token_str[idx] in "rRuUbBfF":
        idx += 1
    prefix = token_str[:idx]
    rest = token_str[idx:]
    if len(rest) >= 6 and rest[:3] in ('"""', "'''") and rest.endswith(rest[:3]):
        quote = rest[:3]
        return prefix, quote, rest[3:-3]
    if len(rest) >= 2 and rest[0] in ('"', "'") and rest.endswith(rest[0]):
        quote = rest[0]
        return prefix, quote, rest[1:-1]
    return None


def _has_theme_c_import(tree: ast.AST) -> bool:
    """Return True if `C` is already imported from `core.ui.theme` (alias `C`)."""
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module == "core.ui.theme":
            for alias in node.names:
                bound_name = alias.asname or alias.name
                if bound_name == "C" and alias.name in ("C", "ThemeColors"):
                    return True
    return False


def _insert_theme_import(source: str, newline: str) -> Tuple[str, bool]:
    """Insert `from core.ui.theme import C` cleanly after existing imports."""
    tree = ast.parse(source)
    if _has_theme_c_import(tree):
        return source, False

    lines = source.splitlines(keepends=True)
    insert_after_line = 0  # 1-based line number after which import goes; 0 = top of file

    for stmt in tree.body:
        if isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Constant) and isinstance(stmt.value.value, str):
            insert_after_line = max(insert_after_line, stmt.end_lineno or stmt.lineno)
        elif isinstance(stmt, (ast.Import, ast.ImportFrom)):
            insert_after_line = max(insert_after_line, stmt.end_lineno or stmt.lineno)
        elif isinstance(stmt, (ast.Assign, ast.AnnAssign)):
            # Stop once executable/assignment statements start (after docstrings & imports)
            break
        else:
            break

    # If file starts with comment header lines and had no docstring/imports yet, place below comments
    if insert_after_line == 0:
        idx = 0
        while idx < len(lines) and (lines[idx].lstrip().startswith("#") or not lines[idx].strip()):
            idx += 1
        insert_after_line = idx

    import_line = f"from core.ui.theme import C{newline}"
    # Avoid touching the line endings of surrounding lines
    new_lines = lines[:insert_after_line] + [import_line] + lines[insert_after_line:]
    return "".join(new_lines), True


def _apply_span_replacements(
    source: str,
    edits: List[Tuple[Tuple[int, int], Tuple[int, int], str]],
) -> str:
    """Apply non-overlapping `(start, end, replacement)` edits ordered top-to-bottom."""
    if not edits:
        return source

    lines = source.splitlines(keepends=True)
    line_offsets: List[int] = [0]
    for line in lines:
        line_offsets.append(line_offsets[-1] + len(line))

    def to_offset(pos: Tuple[int, int]) -> int:
        row, col = pos
        return line_offsets[row - 1] + col

    chars = source
    # Sort descending by start offset so earlier offsets stay valid
    for start_pos, end_pos, repl in sorted(edits, key=lambda e: (e[0][0], e[0][1]), reverse=True):
        s_idx = to_offset(start_pos)
        e_idx = to_offset(end_pos)
        chars = chars[:s_idx] + repl + chars[e_idx:]
    return chars


def _detect_newline(text: str) -> str:
    return "\r\n" if "\r\n" in text else "\n"


def migrate_source(
    source: str,
    token_map: Optional[Dict[str, str]] = None,
    add_import: bool = True,
    path: Optional[Path] = None,
) -> MigrationResult:
    """
    Replace canonical `#RRGGBB` hex color literals in Python UI source with `C.<TOKEN>`
    using AST + tokenizer inspection. Preserves comments, docstrings, bytes literals,
    unmapped colors, and QSS braces (`{` / `}` -> `{{` / `}}` when converting to f-string).
    """
    mapping = {k.upper(): v for k, v in (token_map or HEX_TO_TOKEN).items()}
    tree = ast.parse(source)
    doc_spans = _collect_docstring_spans(tree)

    tokens = list(tokenize.generate_tokens(io.StringIO(source).readline))
    fstring_middle_type = getattr(tokenize, "FSTRING_MIDDLE", None)
    fstring_start_type = getattr(tokenize, "FSTRING_START", None)
    fstring_end_type = getattr(tokenize, "FSTRING_END", None)
    ignored_ws = {tokenize.NL, tokenize.NEWLINE, tokenize.INDENT, tokenize.DEDENT, tokenize.COMMENT}
    string_adjacent_types = {
        t for t in (tokenize.STRING, fstring_start_type, fstring_end_type) if t is not None
    }

    # Precompute neighbor string status for implicit concatenation detection
    significant_indices = [i for i, tok in enumerate(tokens) if tok.type not in ignored_ws]
    sig_pos_by_tok_idx = {tok_idx: pos for pos, tok_idx in enumerate(significant_indices)}

    def is_implicitly_concatenated(tok_idx: int) -> bool:
        pos = sig_pos_by_tok_idx[tok_idx]
        if pos > 0 and tokens[significant_indices[pos - 1]].type in string_adjacent_types:
            return True
        if pos + 1 < len(significant_indices) and tokens[significant_indices[pos + 1]].type in string_adjacent_types:
            return True
        return False

    edits: List[Tuple[Tuple[int, int], Tuple[int, int], str]] = []
    replacements = 0
    skipped = 0
    unmapped: Dict[str, int] = {}

    for i, tok in enumerate(tokens):
        if tok.type == tokenize.STRING:
            if _span_in_docstrings(tok.start, tok.end, doc_spans):
                continue
            hex_matches = list(HEX_COLOR_RE.finditer(tok.string))
            if not hex_matches:
                continue

            for m in hex_matches:
                hx = m.group(0).upper()
                if hx not in mapping:
                    unmapped[hx] = unmapped.get(hx, 0) + 1

            if not any(m.group(0).upper() in mapping for m in hex_matches):
                continue

            parts = _split_string_prefix_and_quotes(tok.string)
            if parts is None:
                skipped += 1
                continue
            prefix, quote, body = parts
            if "b" in prefix.lower():
                skipped += 1
                continue

            # Standalone exact hex literal -> C.<TOKEN>
            try:
                eval_val = ast.literal_eval(tok.string)
            except Exception:
                eval_val = None

            if (
                isinstance(eval_val, str)
                and len(eval_val) == 7
                and eval_val.upper() in mapping
                and not is_implicitly_concatenated(i)
            ):
                edits.append((tok.start, tok.end, f"C.{mapping[eval_val.upper()]}"))
                replacements += 1
                continue

            # Convert string body into f-string safely by escaping braces first
            new_body = body.replace("{", "{{").replace("}", "}}")
            count_in_tok = 0

            def _sub_plain(match: re.Match[str]) -> str:
                nonlocal count_in_tok
                hx = match.group(0).upper()
                if hx in mapping:
                    count_in_tok += 1
                    return "{C." + mapping[hx] + "}"
                return match.group(0)

            new_body = HEX_COLOR_RE.sub(_sub_plain, new_body)
            if count_in_tok == 0:
                continue

            clean_prefix = "".join(ch for ch in prefix if ch not in "uU")
            if "f" not in clean_prefix.lower():
                clean_prefix = "f" + clean_prefix
            edits.append((tok.start, tok.end, f"{clean_prefix}{quote}{new_body}{quote}"))
            replacements += count_in_tok

        elif fstring_middle_type is not None and tok.type == fstring_middle_type:
            if _span_in_docstrings(tok.start, tok.end, doc_spans):
                continue
            hex_matches = list(HEX_COLOR_RE.finditer(tok.string))
            if not hex_matches:
                continue

            for m in hex_matches:
                hx = m.group(0).upper()
                if hx not in mapping:
                    unmapped[hx] = unmapped.get(hx, 0) + 1

            # Slice exact characters from source so tokenize's `{{` offset behavior is respected
            lines = source.splitlines(keepends=True)
            if tok.start[0] == tok.end[0]:
                orig_slice = lines[tok.start[0] - 1][tok.start[1] : tok.end[1]]
            else:
                chunks = [lines[tok.start[0] - 1][tok.start[1] :]]
                for ln in range(tok.start[0] + 1, tok.end[0]):
                    chunks.append(lines[ln - 1])
                chunks.append(lines[tok.end[0] - 1][: tok.end[1]])
                orig_slice = "".join(chunks)

            count_in_tok = 0

            def _sub_fstr(match: re.Match[str]) -> str:
                nonlocal count_in_tok
                hx = match.group(0).upper()
                if hx in mapping:
                    count_in_tok += 1
                    return "{C." + mapping[hx] + "}"
                return match.group(0)

            new_slice = HEX_COLOR_RE.sub(_sub_fstr, orig_slice)
            if count_in_tok > 0:
                edits.append((tok.start, tok.end, new_slice))
                replacements += count_in_tok

    if not edits:
        return MigrationResult(
            path=path,
            original=source,
            updated=source,
            replacements=0,
            unmapped_hexes=unmapped,
            skipped_tokens=skipped,
            import_added=False,
        )

    updated = _apply_span_replacements(source, edits)
    import_added = False
    if add_import:
        updated, import_added = _insert_theme_import(updated, _detect_newline(source))

    # Safety check: verify resulting source compiles cleanly as valid Python AST
    ast.parse(updated)

    return MigrationResult(
        path=path,
        original=source,
        updated=updated,
        replacements=replacements,
        unmapped_hexes=unmapped,
        skipped_tokens=skipped,
        import_added=import_added,
    )


def migrate_file(
    path: Path,
    apply: bool = False,
    token_map: Optional[Dict[str, str]] = None,
) -> MigrationResult:
    """Run migration on `path` (`apply=False` keeps disk untouched)."""
    original = path.read_text(encoding="utf-8")
    result = migrate_source(original, token_map=token_map, add_import=True, path=path)
    if apply and result.changed:
        with path.open("w", encoding="utf-8", newline="") as fh:
            fh.write(result.updated)
    return result


def collect_target_files(repo_root: Path, paths: Sequence[str]) -> List[Path]:
    """Resolve user-supplied paths or default to all UI Python modules under `<repo_root>/src`."""
    if paths:
        resolved: List[Path] = []
        for raw in paths:
            p = Path(raw)
            if not p.is_absolute():
                p = repo_root / p
            if p.is_dir():
                for child in sorted(p.rglob("*.py")):
                    rel = _to_repo_posix(child, repo_root)
                    txt = child.read_text(encoding="utf-8", errors="replace")
                    if is_ui_python_file(rel, text=txt):
                        resolved.append(child)
            elif p.is_file():
                resolved.append(p)
        return resolved

    targets: List[Path] = []
    for p in sorted((repo_root / "src").rglob("*.py")):
        rel = _to_repo_posix(p, repo_root)
        txt = p.read_text(encoding="utf-8", errors="replace")
        if is_ui_python_file(rel, text=txt):
            targets.append(p)
    return targets


def format_unified_diff(result: MigrationResult, repo_root: Path) -> str:
    """Format a unified diff for a changed `MigrationResult`."""
    if not result.changed:
        return ""
    label = _to_repo_posix(result.path, repo_root) if result.path else "<memory>"
    diff_lines = difflib.unified_diff(
        result.original.splitlines(keepends=True),
        result.updated.splitlines(keepends=True),
        fromfile=f"a/{label}",
        tofile=f"b/{label}",
    )
    return "".join(diff_lines)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Migrate raw `#RRGGBB` hex literals in 3GPP Tools UI files to `core.ui.theme.C` tokens."
    )
    parser.add_argument(
        "paths",
        nargs="*",
        help="Optional files or directories to migrate (defaults to all UI Python files under src/).",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="Preview changes without writing files (default).")
    mode.add_argument("--apply", action="store_true", help="Write migrated files back to disk.")
    parser.add_argument("--diff", action="store_true", help="Print unified diffs for modified files.")
    parser.add_argument("--top", type=int, default=0, help="Limit printed diff/summary output to top N files.")
    args = parser.parse_args(argv)

    repo_root = Path(__file__).resolve().parents[3]
    targets = collect_target_files(repo_root, args.paths)
    results: List[MigrationResult] = []
    total_replacements = 0

    for target in targets:
        res = migrate_file(target, apply=args.apply)
        if res.replacements > 0 or res.unmapped_hexes:
            results.append(res)
            total_replacements += res.replacements

    changed_results = sorted([r for r in results if r.changed], key=lambda r: -r.replacements)
    mode_label = "APPLIED" if args.apply else "DRY-RUN"
    print(f"[{mode_label}] Scanned {len(targets)} UI file(s): {len(changed_results)} file(s) modifiable, {total_replacements} hex literal(s) mapped to tokens.")

    shown = changed_results[: args.top] if args.top > 0 else changed_results
    for r in shown:
        rel = _to_repo_posix(r.path, repo_root) if r.path else "<memory>"
        unmapped_cnt = sum(r.unmapped_hexes.values())
        imp = " +import" if r.import_added else ""
        print(f"  {r.replacements:4d} replaced  ({unmapped_cnt:3d} unmapped){imp}  {rel}")
        if args.diff:
            diff_text = format_unified_diff(r, repo_root)
            if diff_text:
                print(diff_text.rstrip())

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
