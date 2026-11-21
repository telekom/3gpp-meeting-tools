# Topic Dossier Exporter Package (`topic_dossier/`)

Modular decomposition of `TopicDossierExporterThread` and standalone Stage 2 checkup helpers previously housed in `topic_dossier_exporter.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`exporter_thread.py`](exporter_thread.py) | `QThread` entry point, signals, cancellation, mixin composition | `TopicDossierExporterThread`, `progress_msg`, `progress_step`, `session_finished`, `cancel`, `is_cancelled`, `run` |
| [`path_utils.py`](path_utils.py) | TDoc ID normalization, path/folder sanitization, media copy, checkpoint keys, shim forwarding | `_normalize_tdoc_id`, `sanitize_folder_name`, `safe_relpath`, `safe_markdown_rel_link`, `get_default_topic_export_base`, `_resolve_doc_server_url`, `_decision_checkpoint_key`, `_copy_markdown_with_media`, `call_convert_document_to_markdown` |
| [`stage2_discovery.py`](stage2_discovery.py) | Stage 2 synthesis package scanning and freshness evaluation | `_read_json_dict`, `_extract_report_header_meta`, `_parse_stage2_subfolder`, `discover_stage2_packages`, `_evaluate_stage2_synthesis_state` |
| [`checkup.py`](checkup.py) | Non-destructive incremental checkup & delta rebuild engine | `_resolve_topic_for_checkup`, `_rebuild_meeting_state_for_checkup`, `check_and_refresh_dossier` |
| [`mixins/export_pipeline.py`](mixins/export_pipeline.py) | Chronological cohort pipeline orchestrator | `TopicDossierExportPipelineMixin._run_export` |
| [`mixins/cache_and_conversion.py`](mixins/cache_and_conversion.py) | Revision lineage, shared canonical Markdown cache, FTP download & Word COM conversion | `TopicDossierCacheAndConversionMixin` (`_build_revision_lineage_map`, `_backfill_canonical_cache`, `_restore_from_canonical_cache`, `_find_or_download_tdoc_file`, `_convert_docs_pipeline`) |
| [`mixins/discovery.py`](mixins/discovery.py) | Topic rule evaluation, TDocs/Chair's Notes loading, live meeting refresh & pruning | `TopicDossierDiscoveryMixin` (`_matches_topic_rule`, `_is_meeting_live_for_refresh`, `_load_meeting_tdocs`, `_load_meeting_agreements`) |
| [`mixins/artifacts.py`](mixins/artifacts.py) | Agreement Markdown export, LLM delta feed, master index, consolidated briefing | `TopicDossierArtifactsMixin` (`_write_agreement_markdown`, `_write_agreements_timeline`, `_write_llm_delta_feed`, `_write_master_index`, `_write_consolidated_briefing`) |
| [`__init__.py`](__init__.py) | Package exports | `TopicDossierExporterThread`, `check_and_refresh_dossier`, `discover_stage2_packages`, path helpers |

## Backward Compatibility & Shim Monkeypatching

- [`../topic_dossier_exporter.py`](../topic_dossier_exporter.py) remains as a thin re-export shim exposing all public and internal symbols used by [`topic_dossier_dialog.py`](../../../ui/dialogs/topic_dossier_dialog.py), [`topic_inputs_dialog.py`](../../../ui/dialogs/topic_inputs_dialog.py), and the test suite.
- `TopicDossierCacheAndConversionMixin` dispatches document conversions through `call_convert_document_to_markdown()`, which checks `sys.modules["modules.meetings.core.exporters.topic_dossier_exporter"].convert_document_to_markdown` so test patches on the legacy shim work transparently.
