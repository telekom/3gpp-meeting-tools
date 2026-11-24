# `modules.specifications.ui.spec_tdoc_report`

Modular mixin decomposition of `SpecTDocReportDialog`, backing the legacy `modules.specifications.ui.spec_tdoc_report_dialog` shim.

## Architecture

- `mixins/ui_setup.py` — `SpecTDocReportUiSetupMixin`: header KPI ribbon, search/facet filters, contribution table, bottom detail drawer, and Markdown dossier tab.
- `mixins/filtering_and_table.py` — `SpecTDocReportTableMixin`: facet combo population, text/combo filtering, table population, and detail drawer HTML rendering.
- `mixins/actions_and_export.py` — `SpecTDocReportActionsMixin`: double-click/menu Word launching, portal/citation actions, and clipboard/Markdown/CSV export.
- `dialog.py` — `SpecTDocReportDialog` preserving modeless top-level `Qt.Window`/`Qt.NonModal` flags and parent lifetime cleanup.
- `__init__.py` — Public package exports.
