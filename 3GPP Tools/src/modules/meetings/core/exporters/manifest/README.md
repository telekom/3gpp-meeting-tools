# Export Manifest & Changelog Package (`exporters/manifest`)

Modular subsystem for `.export_manifest.json` state tracking, incremental session planning, meeting snapshot diffing, legacy reconstruction, and Markdown delta changelog generation.

## Acyclic Module Dependency Graph

```text
validation.py (constants, JSON helpers, folder sanitization, is_export_complete)
    └── models.py (IncrementalPlan, ExportManifest)
            ├── persistence.py (get/load/save manifest, legacy directory reconstruction)
            ├── planning.py (plan_incremental_session, merge_manifest_into_tree)
            └── changelog.py (master & local WI <details> changelogs)

snapshots.py (extract_base_tdoc_id, compute_meeting_snapshot, diff_meeting_snapshots)
```

- `validation.py` and `snapshots.py` are pure leaf modules with zero intra-package dependencies.
- `models.py` depends only on `.validation`.
- `persistence.py`, `planning.py`, and `changelog.py` depend strictly downward on `.models` and `.validation`.
- `export_manifest.py` remains as a thin re-export shim preserving 100% import compatibility (including lazy exports in `exporters/__init__.py`).

## Module Inventory

| File | Responsibility |
|---|---|
| `validation.py` | `MANIFEST_FILENAME`, `MANIFEST_VERSION`, `_clean_for_json`, `_json_fallback`, `sanitize_folder_name`, `is_export_complete` |
| `models.py` | `IncrementalPlan`, `ExportManifest` (serialization, checkpoints, session recording) |
| `snapshots.py` | `extract_base_tdoc_id`, `compute_meeting_snapshot`, `diff_meeting_snapshots` (SHA-1 meeting fingerprints & deltas) |
| `persistence.py` | `get_manifest_path`, `load_manifest`, `save_manifest`, `reconstruct_manifest_from_directory`, `load_or_reconstruct_manifest` |
| `planning.py` | `plan_incremental_session`, `merge_manifest_into_tree` |
| `changelog.py` | `format_master_changelog_markdown`, `format_local_wi_changelog_markdown` |

## Key Invariants

- **Atomic Manifest Writes:** `save_manifest()` writes to `.export_manifest.json.tmp` via `to_extended_path()` and promotes atomically with `os.replace()`.
- **Export Completeness Validation (`is_export_complete`):** Rejects missing/0-byte files, `< 50`-byte stubs, `< 120`-byte files without converter markers, reserved/unavailable placeholders, `"__-________"` corrupt prefixes, presentation exports lacking `"## Slide "`, and Markdown files referencing missing local/media slide PNGs.
- **Non-Destructive Resumability:** `plan_incremental_session()` and `merge_manifest_into_tree()` preserve prior valid exports cleanly across narrower or wider filter runs.
