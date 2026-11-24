# TDocs Agenda & Chairman's Notes Controller (`ui/tdocs_agenda`)

Acyclic modular breakdown behind the backward-compatible `modules.meetings.ui.tdocs_agenda_controller` shim:

- **`refresh_state.py`**: Monotonic sequence guard, `refresh_btn` busy/progress/unlock helpers, and `_facade_attr` indirection so `QMessageBox` / `QTimer` test patches on `tdocs_agenda_controller` apply seamlessly.
- **`agenda_import.py`**: Batch Word/ZIP/HTML agenda ingestion queue (`import_word_agenda_dialog`, `process_imported_agenda_files`) and official Agenda CSV download actions.
- **`chair_notes_and_caching.py`**: Chairman's Notes directory probing/picker/downloading, local RAN notes scanning, and parallel `Docs/` TDoc batch caching.
