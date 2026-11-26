# Email AI Summary Engine (`modules/emails/core/ai_summary/`)

Incremental multi-scope email summarization persistence and 2-stage Per-AI Map-Reduce orchestration.

## Modules

- `models.py`: `SummaryScope` (`thread`, `tdoc`, `ai`, `selection`), `EmailSummaryRecord`, and `SummaryFreshness`.
- `store.py`: `EmailSummaryStore` (`<meeting_dir>/Agenda/email_summaries.db`) with `cutoff_utc`, `covered_message_ids` set-difference freshness, `incremental_hops`, and 5-deep undo history.
- `formatting.py`: Preserves `<<START>>..<<END>>` blocks via `clean_body`, extracts deterministic facts, formats `[Msg #N]` blocks, and parses optional ```` ```summary-meta ```` blocks fault-tolerantly.
- `orchestrator.py`: `EmailSummaryOrchestrator` (pure Python) supporting incremental delta merges, ad-hoc selection comparisons, and 2-stage Map-Reduce Per-AI cross-thinking.
- `workers.py`: `EmailSummaryTaskWorker` and `AiBriefingWorker` (`QThread` wrappers with `CancellationFlag`).
