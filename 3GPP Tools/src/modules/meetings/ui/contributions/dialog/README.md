# `modules.meetings.ui.contributions.dialog`

Modular modeless cross-meeting Company Contribution Search & Report dialog (`ContributionReportDialog`).

## Package Architecture (Acyclic)

- `mixins/ui_setup.py`: `ContributionDialogUiSetupMixin` (filter sidebar, KPI cards, results table, column widths, 100% `core.ui.theme.C` tokens).
- `mixins/search_and_kpis.py`: `ContributionDialogSearchMixin` (`ContributionSearchWorker` orchestration, filter summary, KPI aggregation).
- `mixins/row_actions_and_session.py`: `ContributionDialogActionsMixin` (table double-clicks, context menus, TDoc downloads, Excel/HTML stats delegation, session save/restore).
- `report_dialog.py`: `ContributionReportDialog` orchestrator combining the mixins and cleanly terminating background threads.
- `__init__.py`: Public facade preserving all 12 compatibility exports.

## Compatibility

`src/modules/meetings/ui/contributions/contribution_dialog.py` remains a thin backward-compatible re-export shim.
