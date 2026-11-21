# Topic Digest & Briefing Builder (`exporters/topic_digest/`)

Acyclic modular decomposition of the deterministic per-meeting briefing and LLM digest generator
previously housed in the 1,541-line `topic_digest_builder.py`. `topic_digest_builder.py` remains
a backward-compatible re-export shim exposing all 30 public/helper symbols.

## Submodule Navigation Map (Acyclic DAG)

| Layer / Module | Responsibility | Key Symbols |
|---|---|---|
| [`constants.py`](constants.py) | Stage-1/Stage-2 filenames, delta markers, size budgets & status regexes | `STAGE1_DELTA_START`, `STAGE1_DELTA_END`, `STAGE2_SYNTHESIS_FILENAME`, `STAGE2_REPORT_FILENAME`, `DEFAULT_MAX_DIGEST_CHARS`, `_SUFFIX_IN_NUM_RE`, `_TREATED_STATUS_TOKENS` |
| [`extraction.py`](extraction.py) | Meeting folder IDs, media rewriting, contributor matrices & TDoc position digesting | `resolve_meeting_folder_id`, `rewrite_media_links`, `_rel_to_meeting_dir`, `build_contributor_matrix`, `_strip_markdown_frontmatter`, `extract_tdoc_position_digest` |
| [`status_and_outcomes.py`](status_and_outcomes.py) | Lifecycle/analysis badges, Stage-2 synthesis state, TDoc outcome reconciliation & digest cards | `classify_meeting_lifecycle`, `format_lifecycle_badge`, `format_analysis_badge`, `format_stage2_summary_badge`, `build_tdoc_outcome_index`, `_format_tdoc_digest_card` |
| [`meeting_brief.py`](meeting_brief.py) | Per-meeting `_meeting_brief.md` generation & non-destructive Stage-1 delta merging | `build_delta_section_markdown`, `sync_meeting_brief_analysis_state`, `write_meeting_brief` |
| [`company_positions.py`](company_positions.py) | Token-bounded Stage-1 `_company_positions_input*.md` packets partitioned by company | `write_company_positions_input` |
| [`root_reports.py`](root_reports.py) | Root E2E prompt/checklist, master `index.md`, and `consolidated_briefing.md` | `write_e2e_synthesis_prompt`, `write_master_index_markdown`, `write_consolidated_briefing_markdown` |

## Invariants
- **Strict Acyclic Dependency Order:** `constants` $\rightarrow$ `extraction` $\rightarrow$ `status_and_outcomes` $\rightarrow$ `{meeting_brief, company_positions, root_reports}` $\rightarrow$ `__init__`.
- **Non-Destructive Stage-1 Delta Rule (Rule 7 in `meetings/AGENTS.md`):** `sync_meeting_brief_analysis_state()` injects/clears `STAGE1_DELTA_START`..`STAGE1_DELTA_END` in-place inside `_meeting_brief.md` without overwriting `_company_positions_analysis.md` or emitting `.delta.md` sibling files; Stage-2 prompt targets dated `stage2_<YYYY-MM-DD>_rN/` packages.
- Keep every file under **510 lines**.
