# Chairman's Notes Downloader Package (`core/chair_notes`)

Acyclic modular breakdown behind the backward-compatible `modules.meetings.core.chair_notes_downloader` shim:

- **`selector.py`**: Filename eligibility, HTML directory listing parsing, clean-vs-tracked deduplication, and latest-by-track selection.
- **`transfer.py`**: Subpath probing, atomic `.part` HTTP streaming downloads, progress byte formatting, and `.zip` extraction.
- **`threads.py`**: `ChairNotesListingThread` and `ChairNotesDownloaderThread` workers preserving all static/class delegates.
