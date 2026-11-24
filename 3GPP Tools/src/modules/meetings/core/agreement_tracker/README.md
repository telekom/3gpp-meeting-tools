# `modules.meetings.core.agreement_tracker`

Modular cross-meeting Study Item & Topic agreement scanner (`AgreementTrackerWorker`).

## Package Architecture (Acyclic)

- `filters.py`: `AgreementFilterMixin` (AI/WI/keyword matching, sorting keys, `agenda.csv` mapping, candidate FTP URL assembly).
- `fetchers.py`: `AgreementFetchersMixin` (TDocs Excel workbook and Chair Notes cache/FTP download/parse, `.doc` -> `.docx` LibreOffice fallback).
- `processor.py`: `AgreementProcessorMixin` (`_process_single_meeting` enriching Chair Notes agreements and Agreed Excel TDocs).
- `worker.py`: `AgreementTrackerWorker` QThread orchestrator running the bounded pool and emitting progress/results.
- `__init__.py`: Public facade exporting `AgreementTrackerWorker` and `logger`.

## Compatibility

`src/modules/meetings/core/agreement_tracker_worker.py` remains a thin backward-compatible re-export shim.
