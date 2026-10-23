# --- File: src/modules/meetings/core/stats/plot_status.py ---
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from core.config.plot_styles import STATUS_COLORS, STATUS_ORDER, DOC_TYPE_COLORS, normalize_status
from .stats_html import dataframe_to_html_table


def generate_outcomes_plot(df, export_dir, palette=None, prefix_id="Global", save_html=False):
    """Generates a donut pie chart showing overall contribution outcomes with canonical semantic colors."""
    status_col = 'Clean_Status' if 'Clean_Status' in df.columns else 'TDoc Status'
    work_df = df.copy()
    if status_col not in work_df.columns:
        work_df[status_col] = 'Unknown'
    work_df[status_col] = work_df[status_col].apply(normalize_status)

    status_counts = work_df[status_col].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    status_counts = status_counts[status_counts['Status'].astype(str).str.strip() != '']

    # Sort according to canonical STATUS_ORDER
    status_counts['order_rank'] = status_counts['Status'].apply(
        lambda s: STATUS_ORDER.index(s) if s in STATUS_ORDER else len(STATUS_ORDER)
    )
    status_counts = status_counts.sort_values('order_rank').drop(columns=['order_rank']).reset_index(drop=True)

    total = int(status_counts['Count'].sum())
    status_counts['Percentage'] = status_counts['Count'].apply(lambda x: f"{(100*x/total):.1f}%" if total else "0.0%")

    ordered_labels = status_counts['Status'].tolist()
    ordered_values = status_counts['Count'].tolist()
    ordered_colors = [STATUS_COLORS.get(s, '#94A3B8') for s in ordered_labels]

    fig_status = go.Figure(data=[go.Pie(
        labels=ordered_labels,
        values=ordered_values,
        hole=0.42,
        sort=False,
        marker=dict(colors=ordered_colors, line=dict(color='#FFFFFF', width=1.5)),
        textinfo='percent+label',
        hovertemplate="<b>%{label}</b><br>TDocs: %{value}<br>Percentage: %{percent}<extra></extra>",
        customdata=ordered_labels
    )])
    fig_status.update_traces(meta={'filterDimension': 'status'})
    fig_status.update_layout(
        title="TDoc Outcomes",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        showlegend=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02)
    )

    if save_html:
        fig_status.write_html(str(export_dir / f"{prefix_id}_Outcomes.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Outcomes'}}
    table = status_counts.rename(columns={'Count': 'TDocs'})
    return (fig_status.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                               config=svg_config), dataframe_to_html_table(table))


def generate_doc_types_plot(df, export_dir, prefix_id="Global", save_html=False):
    """Generates a stacked horizontal bar chart showing Document Types segmented by outcome."""
    type_col = 'Clean_Doc_Type' if 'Clean_Doc_Type' in df.columns else 'Type'
    status_col = 'Clean_Status' if 'Clean_Status' in df.columns else 'TDoc Status'

    work_df = df.copy()
    if status_col not in work_df.columns:
        work_df[status_col] = 'Unknown'
    work_df[status_col] = work_df[status_col].apply(normalize_status)

    valid_df = work_df[(work_df[type_col].astype(str).str.strip() != '') &
                       (work_df[status_col].astype(str).str.strip() != '')].copy()

    if valid_df.empty:
        msg = "<p style='padding:20px; color:#666;'>No document type data available.</p>"
        return msg, "<p class='data-empty'>No data available.</p>"

    counts = valid_df.groupby([type_col, status_col]).size().reset_index(name='Count')
    totals = valid_df[type_col].value_counts()
    doc_types = totals.index.tolist()

    # Reorder doc_types ascending so the most frequent appears at top of horizontal stacked bar
    doc_types_ascending = list(reversed(doc_types))

    fig = go.Figure()
    for status in STATUS_ORDER:
        status_data = counts[counts[status_col] == status]
        if not status_data.empty:
            type_to_count = dict(zip(status_data[type_col], status_data['Count']))
            x_vals = [int(type_to_count.get(t, 0)) for t in doc_types_ascending]
            if sum(x_vals) > 0:
                fig.add_trace(go.Bar(
                    name=status,
                    y=doc_types_ascending,
                    x=x_vals,
                    orientation='h',
                    marker=dict(color=STATUS_COLORS.get(status, '#94A3B8')),
                    hovertemplate=f"<b>%{{y}}</b><br>Outcome: {status}<br>TDocs: %{{x}}<extra></extra>",
                    customdata=doc_types_ascending
                ))

    fig.update_layout(
        barmode='stack',
        title="Document Types by Outcome",
        xaxis_title="Contributions Count",
        yaxis_title=None,
        legend_title_text="Outcome",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02)
    )

    if save_html:
        fig.write_html(str(export_dir / f"{prefix_id}_Doc_Types.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Doc_Types'}}

    # Tabular summary with status breakdown and agreement rate
    pivot = valid_df.pivot_table(index=type_col, columns=status_col, aggfunc='size', fill_value=0)
    ordered_cols = [s for s in STATUS_ORDER if s in pivot.columns] + [c for c in pivot.columns if c not in STATUS_ORDER]
    pivot = pivot[ordered_cols]
    pivot['Total'] = pivot.sum(axis=1)

    agreed_cols = [c for c in ['Agreed', 'Approved'] if c in pivot.columns]
    if agreed_cols:
        pivot['Agreed %'] = ((pivot[agreed_cols].sum(axis=1) / pivot['Total']) * 100).round(1).astype(str) + '%'
    else:
        pivot['Agreed %'] = '0.0%'

    table_df = pivot.sort_values('Total', ascending=False).reset_index().rename(columns={type_col: 'Document Type'})

    return (fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                        config=svg_config), dataframe_to_html_table(table_df))
