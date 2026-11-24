# `modules.specifications.core.db`

Modular SQLite catalogue package for [`SpecsDatabase`](database.py), decomposed from the former `database.py` monolith:

- [`schema.py`](schema.py): Thread-safe SQLite connection context manager, DDL/WAL setup, orphan garbage collection, vacuum, and filter dropdown options.
- [`ingestion.py`](ingestion.py): File/version upsert, portal upload-date sync, Work Item linking, DynaReport metadata updates, manual spec upsert, deletion, and wipe.
- [`queries.py`](queries.py): Filtered file/spec search (`search_files`, `get_filtered_specs`), title maps, specification details, archive version sorting, and related WIs.
- [`annotations.py`](annotations.py): Per-specification user tags and comments CRUD.
- [`database.py`](database.py): `SpecsDatabase` class composing the four mixins.
