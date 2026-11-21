# `tdocs/models` Package

Modular implementation of `TDocsTableModel`, `TDocsFilterProxyModel`, and `natural_sort_key`, decomposed cleanly from `src/modules/meetings/ui/tdocs_models.py`.

## Structure & Mixin Composition (MRO)

- `sorting.py`: `@lru_cache` `natural_sort_key()` for sub-millisecond Agenda/TDoc comparisons.
- `proxy_model.py` (`TDocsFilterProxyModel`): Multi-facet filtering (`rowMatchesFilters`, boolean queries, tracked WIs, unassigned LS in, tentative company toggle) and composite sorting (`lessThan`).
- `table_model.py` (`TDocsTableModel`): `QAbstractTableModel` shell combining five mixins ahead of `QAbstractTableModel`:
  1. `TDocsIndexingMixin` (`mixins/indexing.py`): WI token precomputation, `_agenda_sort_tuple`, company sanitization, revision lineage cache, email stats.
  2. `TDocsStateThreadsMixin` (`mixins/state_and_threads.py`): User annotation inheritance, tracked email thread row injection (`sync_tracked_thread_rows`), loading/highlight states.
  3. `TDocsAgreementsHtmlMixin` (`mixins/agreements_and_html.py`): Hyperlink formatting, Chairman's Notes agreement tagging, rich HTML remarks/tooltips.
  4. `TDocsDataRolesMixin` (`mixins/data_roles.py`): `QAbstractTableModel.data()` implementation across display/tooltip/background/foreground/sort roles.
  5. `TDocsAgendaMergeMixin` (`mixins/agenda_merge.py`): Live `TDocsByAgenda` / Chair's Notes merge and on-the-fly revision injection.

## Backward Compatibility

`src/modules/meetings/ui/tdocs_models.py` acts as a thin re-export shim exposing `natural_sort_key`, `TDocsTableModel`, `TDocsFilterProxyModel`, and `logger`.
