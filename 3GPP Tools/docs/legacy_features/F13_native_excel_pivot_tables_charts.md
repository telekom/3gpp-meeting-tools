# F-13: Native Excel COM Pivot Tables & Charts

## 1. Executive Summary & Objective

In standards delegate workflows, delegates frequently need statistical breakdowns of meeting activity, such as:
* Document volume per Agenda Item (`TDocs per AI`).
* Document distribution by category (`CR`, `Discussion`, `LS`, `Work Item Description`).
* Most active companies and co-sponsorship matrices.

In modern **3GPP Tools**, this need is addressed through responsive HTML dashboards (`Statistics_Report.html`) powered by Plotly and Chart.js, accompanied by static multi-tab Excel workbooks generated via `openpyxl`.

However, the legacy application utilized Microsoft Excel COM automation (`win32com.client`) to construct **dynamic, interactive native Excel Pivot Tables and Pivot Charts** directly within exported `.xlsx` workbooks. These dynamic pivot tables allowed delegates working inside Excel to slice, filter, drill down, and reorganize meeting data using Excel's native PivotTable Field List.

The objective of **F-13** is to document this legacy COM pivot table automation, analyze its viability, and evaluate whether native Excel Pivot creation should be offered alongside modern HTML dashboards.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/parsing/excel.py` (functions: `add_pivot_table_to_workbook_active_sheet`, `generate_pivot_chart_from_tdocs_by_agenda`)

### Mechanism & Implementation Details

#### Native PivotTable Creation via Excel COM
The legacy script used Excel's VBA COM object model:

```python
# parsing/excel.py
def add_pivot_table_to_workbook_active_sheet(
        wb, cell_range, data_field=None, row_field=None,
        column_field=None, ws_name='', order_by=''):

    ws = wb.ActiveSheet
    pivot_table_source_data = ws.Range(cell_range)
    
    # XlPivotTableSourceType: xlDatabase = 1
    # XlPivotTableVersion: xlPivotTableVersion15 (Excel 2013+) = 5
    pivot_cache = wb.PivotCaches().Create(1, pivot_table_source_data, 5)

    last_wb = wb.Sheets(wb.Sheets.Count)
    pivot_chart_wb = wb.Sheets.Add(Before=None, After=last_wb)
    pivot_chart_wb.Name = ws_name

    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=pivot_chart_wb.Range("A1"),
        TableName=f'{ws_name} Pivot Table'
    )

    # Configure Pivot Table Properties
    pivot_table.ColumnGrand = True
    pivot_table.RowGrand = True
    pivot_table.EnableDrilldown = True

    # Pivot Field Orientations:
    # xlRowField = 1, xlColumnField = 2, xlDataField = 4
    if row_field:
        pivot_table.PivotFields(row_field).Orientation = 1
    if column_field:
        pivot_table.PivotFields(column_field).Orientation = 2
    if data_field:
        pivot_table.PivotFields(data_field).Orientation = 4

    # Auto-sort
    if order_by:
        pivot_table.PivotFields(order_by).AutoSort(2) # xlDescending = 2
```

#### Native Pivot Chart Generation (`generate_pivot_chart_from_tdocs_by_agenda`)
```python
def generate_pivot_chart_from_tdocs_by_agenda(wb):
    summary_chart = wb.Charts.Add()
    summary_chart.Name = "TDocs per AI"
    summary_chart.ChartType = 52  # xlColumnStacked
    summary_chart.HasTitle = False
    summary_chart.HasLegend = True
    summary_chart.ShowAllFieldButtons = False
    
    x_axis = summary_chart.Axes(1)
    x_axis.HasTitle = True
    x_axis.AxisTitle.Text = "Agenda Item"
    
    y_axis = summary_chart.Axes(2)
    y_axis.HasTitle = True
    y_axis.AxisTitle.Text = "TDocs"
```

---

## 3. Current State in 3GPP Tools & Comparison

| Aspect | Legacy Approach (F-13) | Modern Approach (3GPP Tools) |
| :--- | :--- | :--- |
| **Technology** | Win32 COM (`Excel.Application`) | `openpyxl` + Plotly HTML (`Statistics_Report.html`) |
| **Portability** | Windows + Desktop Excel only | Cross-platform (runs anywhere Python runs) |
| **Interactivity** | Native Excel PivotTable fields | Rich web graphs (zoom, pan, hover tooltips, alliance networks) |
| **Execution Speed** | Slow (2–5 seconds, requires Excel launch) | Sub-second generation via openpyxl/HTML |
| **Offline Presentation**| Requires Excel license | Opens in any web browser without Office installed |

---

## 4. Architectural Assessment & Recommendations

### Assessment
* For 90% of reporting and debriefing workflows, the **modern HTML analytics dashboard (`Statistics_Report.html`) is substantially superior** to Excel COM pivot charts in aesthetics, shareability, and functionality (providing Louvain alliance clustering, responsive KPI strips, and interactive toggle tables).
* However, for delegates who perform custom deep-dive modeling inside Microsoft Excel, having **dynamic PivotTables** already configured in the generated `.xlsx` workbook (instead of static data tabs) remains a nice-to-have utility.

### Modern Implementation Options
1. **Option A: Headless OpenXML Pivot Tables via `openpyxl`:**
   * Modern `openpyxl` has rudimentary support for PivotTables without requiring Excel COM.
2. **Option B: Post-Processing Script via Windows COM:**
   * An optional `"Enrich Excel with Dynamic Pivot Tables"` action under `Export ▾` that triggers COM post-processing on Windows machines with Excel installed.

---

## 5. Implementation Roadmap & Acceptance Criteria

### Priority Rating
* **Priority:** ⚪ Low (Deprioritized relative to F-01, F-02, F-03, F-04, and F-12, because HTML dashboards satisfy the primary use case).

### Future Implementation Tasks
1. [ ] If requested by delegates, add an optional COM post-step to `src/modules/meetings/core/excel_exporter.py` that injects native PivotCaches.
2. [ ] Add toggle in `MeetingsConfigDialog`: `[ ] Generate Excel COM Pivot Tables on export (requires Excel)`.
