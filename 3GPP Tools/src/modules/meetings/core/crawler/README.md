# `modules.meetings.core.crawler`

Acyclic package backing the legacy `modules.meetings.core.scraper` shim.

## Architecture

- `sources.py` — Working group FTP/DynaReport endpoints (`MEETING_SOURCES`), regex patterns, deny list, and pure folder-name helpers (`is_meeting_name`, `extract_meeting_number`).
- `ftp_and_docs.py` — WG FTP directory listing (`fetch_wg_directories`) and `Docs/` TDoc range scanner (`process_individual_meeting`).
- `dynareport.py` — DynaReport HTML metadata parser (`process_dynareport`).
- `crawler_thread.py` — 3-phase `MeetingsCrawlerThread(QThread)` syncing WG directories, TDoc bounds, and DynaReports into `MeetingsDatabase`.
- `manual_fetcher.py` — Single-meeting `ManualMeetingFetcherThread(QThread)` resolver for user queries and direct FTP URLs.
- `__init__.py` — Public package exports (`MEETING_SOURCES`, `MeetingsCrawlerThread`, `ManualMeetingFetcherThread`).

## Dependency Flow

`sources` -> `ftp_and_docs` -> `crawler_thread` -> `manual_fetcher` (`dynareport` -> `crawler_thread`).
