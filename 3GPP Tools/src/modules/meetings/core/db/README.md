# `modules.meetings.core.db`

Modular SQLite persistence layer for the 3GPP meetings catalogue (`working_groups` and `meetings` tables in `src/db/3gpp_data.db`).

## Package Architecture (Acyclic)

- `schema_and_writes.py`: WAL connection context manager, idempotent schema creation & `ALTER TABLE` migrations, meeting flag/sort-number helpers, bulk inserts/updates, and single-meeting upsert.
- `queries.py`: Filtered search, working-group listing, deletion helpers, active-sync checks, TDoc-to-meeting resolution (`find_meeting_by_tdoc`, `find_incomplete_tdoc_meetings`), and active/upcoming in-person meeting cohorts (`_NON_PHYSICAL_LOCATIONS`).
- `database.py`: `MeetingsDatabase` orchestrator combining both leaf mixins.
- `__init__.py`: Public re-export of `MeetingsDatabase`.

## Compatibility

`src/modules/meetings/core/meetings_db.py` remains a thin backward-compatible re-export shim for all existing callers and test patch targets.
