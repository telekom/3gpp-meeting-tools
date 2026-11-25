# `modules.meetings.core.crawler`

Acyclic package backing the legacy `modules.meetings.core.scraper` shim.

## Architecture

- `sources.py` — Working group FTP/DynaReport endpoints (`MEETING_SOURCES`), regex patterns, deny list (excluding `TSGR*_AH*` containers), and pure folder-name helpers (`is_meeting_name`, `extract_meeting_number`, `extract_fallback_date_from_folder`, `normalize_meeting_ftp_path`).
- `ftp_and_docs.py` — WG FTP directory listing (`fetch_wg_directories`) and `Docs/` TDoc range scanner (`process_individual_meeting`).
- `dynareport.py` — DynaReport HTML metadata parser (`process_dynareport`) with `Ad Hoc` and root `url_key` normalization.
- `crawler_thread.py` — Fast-first 3-step `MeetingsCrawlerThread(QThread)` syncing (1) WG FTP directories, (2) DynaReport metadata & future meetings, and (3) smart/all/none `Docs/` TDoc bounds with `progress_updated`, `task_summary`, and cooperative `isInterruptionRequested()` cancellation.
- `manual_fetcher.py` — Single-meeting `ManualMeetingFetcherThread(QThread)` resolver for user queries and direct FTP URLs.
- `__init__.py` — Public package exports (`MEETING_SOURCES`, `MeetingsCrawlerThread`, `ManualMeetingFetcherThread`).

## Dependency Flow

`sources` -> `ftp_and_docs` -> `crawler_thread` -> `manual_fetcher` (`dynareport` -> `crawler_thread`).
