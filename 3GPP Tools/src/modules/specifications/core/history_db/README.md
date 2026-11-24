# `modules.specifications.core.history_db`

Modular SQLite engine for `3gpp_spec_history.db` managing Specification Change History, CR/p-CR TDoc cross-references, and TDoc cover-sheet/agenda metadata caches.

## Package Architecture (Acyclic)

- `schema_and_writes.py`: WAL connections, schema initialization & migrations, `upsert_history_records()`, deletion, and `wipe_database()`.
- `history_queries.py`: `get_spec_history()` (`get_history_for_spec`), `search_history()`, version summaries, reverse TDoc lookups, and statistics.
- `tdoc_metadata.py`: `tdoc_metadata_cache` upserts, affected-clause updates, and batched `O(1)` TDoc metadata lookups.
- `database.py`: `SpecHistoryDatabase` orchestrator combining all three leaf mixins.
- `__init__.py`: Public facade exporting `SpecHistoryDatabase` and `logger`.

## Compatibility

`src/modules/specifications/core/spec_history_db.py` remains a thin backward-compatible re-export shim.
