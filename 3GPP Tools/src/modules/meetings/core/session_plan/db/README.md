# Session Plan Database Package (`session_plan/db/`)

Modular decomposition of `SessionPlanDatabase` previously housed in the 1,506-line
`session_plan_db.py`. `session_plan_db.py` remains a backward-compatible re-export shim.

## Where to edit what

| Concern | File |
|---|---|
| `__init__`, global defaults path resolution (test isolation), SQLite WAL connection context manager | `database.py` |
| SQLite DDL (`14` tables + indexes + `room_override` migration), `delegation_defaults.json` sync, `meeting_meta` | `mixins/schema_and_defaults.py` |
| Snapshots CRUD, idempotent `_repair_stale_slot_items()`, user annotations, tracked topics | `mixins/snapshots_and_annotations.py` |
| Delegation roster, default/per-meeting topic owners, session appointments, `DelegationPlan` JSON import/export | `mixins/delegation.py` |
| Physical room & per-slot custom location notes (`meeting_locations.json` bidirectional sync) | `mixins/locations.py` |

## Invariants
- **Test Isolation for `delegation_defaults.json`:** `_resolve_global_defaults_path()` scopes `delegation_defaults.json` to the top-level temporary folder when `db_path` sits inside OS temp, keeping multi-meeting unit tests isolated from production data.
- **SQLite WAL & Concurrency:** `_get_connection()` enables `PRAGMA journal_mode=WAL;` and `PRAGMA busy_timeout=30000;`.
- **Self-Healing Snapshot Items:** `get_snapshot_by_id()` invokes static `_repair_stale_slot_items()` to re-parse stale cached `slot_agenda_items` via `SessionPlanParser._extract_agenda_items()`.
- Keep every file under **490 lines**.
