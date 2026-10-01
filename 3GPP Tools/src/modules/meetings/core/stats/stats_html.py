import html
import json
import pandas as pd


def dataframe_to_html_table(df: pd.DataFrame, empty_message="No data available.") -> str:
    """Render an interactive dashboard table from aggregated plot data.

    Sorting/search/copy/CSV behavior is supplied by the dashboard JavaScript.
    The markup is also harmless when used in a standalone Plotly export.
    """
    if df is None or df.empty:
        return f'<p class="data-empty">{html.escape(empty_message)}</p>'

    safe = df.copy()
    safe.columns = [str(c) for c in safe.columns]
    table_id = f"stats-table-{id(df)}"

    rows = []
    for row_idx, (_, row) in enumerate(safe.iterrows()):
        cells = []
        for col in safe.columns:
            value = row[col]
            if pd.isna(value):
                value = ""
            text = str(value)
            # Keep a raw sortable value separate from the escaped presentation.
            cells.append(
                f'<td data-value="{html.escape(text, quote=True)}">{html.escape(text)}</td>'
            )
        rows.append(f'<tr data-original-index="{row_idx}">' + "".join(cells) + "</tr>")

    headers = "".join(
        f'<th tabindex="0" title="Sort by {html.escape(col)}" '
        f'onclick="sortStatsTable(this)" onkeydown="if(event.key===\'Enter\')sortStatsTable(this)">'
        f'{html.escape(col)} <span class="sort-indicator">↕</span></th>'
        for col in safe.columns
    )
    return (
        '<div class="data-tools">'
        '<input class="data-search" type="search" placeholder="Search table..." '
        'oninput="filterStatsTable(this)" aria-label="Search table">'
        '<span class="data-row-count"></span>'
        '<button type="button" onclick="resetStatsTable(this)">↺ Reset</button>'
        '<button type="button" onclick="copyVisibleTable(this)">📋 Copy</button>'
        '<button type="button" onclick="downloadVisibleTableCsv(this)">⬇ CSV</button>'
        '</div>'
        f'<div class="data-table-wrap"><table class="data-table" id="{table_id}">'
        f'<thead><tr>{headers}</tr></thead><tbody>{"".join(rows)}</tbody>'
        '</table></div>'
    )
