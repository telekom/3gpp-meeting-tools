# Database Maintenance Hub (`core/ui/database_maintenance/`)

Modular decomposition of the SQLite Database Maintenance & Compaction Hub previously housed
in the 809-line `core/ui/database_maintenance_dialog.py`. `database_maintenance_dialog.py`
remains a thin, 100%-compatible re-export shim preserving `DatabaseMaintenanceDialog`,
`KpiCard`, and `format_file_size`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`styles.py`](styles.py) | QSS definitions backed strictly by `core.ui.theme.C` tokens | `DIALOG_BG_STYLE`, `KPI_CARD_STYLE`, `STORAGE_TABLE_STYLE`, `build_kpi_value_style` |
| [`storage_ops.py`](storage_ops.py) | Pure stdlib/SQLite DB discovery, `?mode=ro` freelist metrics, and `VACUUM`/WAL truncation | `format_file_size`, `get_tracked_databases`, `query_db_metrics`, `vacuum_database_file` |
| [`widgets.py`](widgets.py) | Executive KPI summary card widget | `KpiCard` |
| [`table_mixin.py`](table_mixin.py) | 6-column storage table builder & KPI card refresh logic | `_MaintenanceTableMixin` |
| [`actions_mixin.py`](actions_mixin.py) | Folder open, modeless `DatabaseInspectorDialog` launch, single/batch compact, wipe | `_MaintenanceActionsMixin` |
| [`dialog.py`](dialog.py) | Modeless top-level `QDialog` shell, layout setup & delegation hooks | `DatabaseMaintenanceDialog`, `logger` |
| [`__init__.py`](__init__.py) | Package exports | `DatabaseMaintenanceDialog`, `KpiCard`, `format_file_size`, `get_tracked_databases`, `query_db_metrics`, `vacuum_database_file` |

## Invariants
- **Modeless Top-Level Behavior:** `DatabaseMaintenanceDialog` initializes with `super().__init__(None)`, `Qt.Window`, and `Qt.NonModal` (no `Qt.WindowStaysOnTopHint`) to prevent Win32 `GW_OWNER` pinning above external applications.
- **Zero Raw Hex Literals:** All colors and styles reference `core.ui.theme.C` tokens, guarded by `tests/test_ui_hex_lint.py`.
- Keep every file under **250 lines**.
