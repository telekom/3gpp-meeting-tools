# Meeting Statistics Exporter (`modules.meetings.core.stats.exporter`)

Acyclic modular package extracted from `stats/exporter_thread.py` (which remains a thin backward-compatible re-export shim):

- `kpi_and_workbook.py`: `compute_scope_kpis()`, `prepare_statistics_dataframe()`, and `export_statistics_workbook()` (`.xlsx` multi-sheet workbook export).
- `html_shell.py`: Standalone HTML dashboard CSS/JS template and `render_dashboard_html()`.
- `thread.py`: `StatisticsExporterThread(QThread)` orchestrating global, WI, and AI scope plot generation, Excel export, and dashboard output.
