import pandas as pd
import plotly.express as px

from .stats_html import dataframe_to_html_table


def _axis_col(df):
    return 'AI_Axis' if 'AI_Axis' in df.columns else 'Agenda Item'


def generate_ai_volume_plot(df, export_dir, theme_color, prefix_id="Global", save_html=False):
    axis_col = _axis_col(df)
    group_cols = ['Agenda Item', axis_col] if axis_col != 'Agenda Item' else ['Agenda Item']
    ai_counts = df.groupby(group_cols, dropna=False).size().reset_index(name='Count')
    ai_counts = ai_counts[ai_counts['Agenda Item'].astype(str).str.strip() != '']
    ai_counts = ai_counts.sort_values('Count', ascending=False)

    hover_data = {'Agenda Item': True, 'Count': True}
    if 'AI_Full_Description' in df.columns:
        desc = df[['Agenda Item', 'AI_Full_Description']].drop_duplicates('Agenda Item')
        ai_counts = ai_counts.merge(desc, on='Agenda Item', how='left')
        hover_data['AI_Full_Description'] = True

    hover_name = 'AI_Display' if 'AI_Display' in ai_counts.columns else 'Agenda Item'
    if 'AI_Display' in df.columns and 'AI_Display' not in ai_counts.columns:
        meta = df[['Agenda Item', 'AI_Display']].drop_duplicates('Agenda Item')
        ai_counts = ai_counts.merge(meta, on='Agenda Item', how='left')
    fig_ai = px.bar(ai_counts, x=axis_col, y='Count', title="Agenda Items by TDoc Volume",
                    color_discrete_sequence=[theme_color], hover_name=hover_name, hover_data=hover_data,
                    custom_data=['Agenda Item'])
    fig_ai.update_traces(meta={'filterDimension': 'ai'})
    fig_ai.update_xaxes(type='category', categoryorder='total descending', title="Agenda Item / Topic")
    fig_ai.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    if save_html:
        fig_ai.write_html(str(export_dir / f"{prefix_id}_AI_Volume.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_AI_Volume'}}

    table_cols = ['Agenda Item']
    for col in ['AI_Acronym', 'AI_Topic']:
        if col in df.columns:
            meta = df[['Agenda Item', col]].drop_duplicates('Agenda Item')
            ai_counts = ai_counts.merge(meta, on='Agenda Item', how='left')
            table_cols.append(col)
    table_cols.append('Count')
    table = ai_counts[table_cols].rename(columns={'AI_Acronym': 'Acronym', 'AI_Topic': 'Topic', 'Count': 'TDocs'})

    return (fig_ai.to_html(full_html=False, include_plotlyjs='cdn', default_height="100%", default_width="100%",
                           config=svg_config), dataframe_to_html_table(table))


def generate_ai_status_plot(df, export_dir, palette, prefix_id="Global", save_html=False):
    axis_col = _axis_col(df)
    valid_df = df[(df['Agenda Item'].astype(str).str.strip() != '') &
                  (df['TDoc Status'].astype(str).str.strip() != '')].copy()
    counts = valid_df.groupby(['Agenda Item', axis_col, 'TDoc Status']).size().reset_index(name='Count')
    top_ais = valid_df['Agenda Item'].value_counts().index
    plot_df = counts[counts['Agenda Item'].isin(top_ais)]

    if 'AI_Display' in valid_df.columns and 'AI_Display' not in plot_df.columns:
        meta = valid_df[['Agenda Item', 'AI_Display']].drop_duplicates('Agenda Item')
        plot_df = plot_df.merge(meta, on='Agenda Item', how='left')
    hover_name = 'AI_Display' if 'AI_Display' in plot_df.columns else 'Agenda Item'
    fig = px.bar(plot_df, x=axis_col, y='Count', color='TDoc Status', title="Agenda Items by Outcome Status",
                 color_discrete_sequence=palette, barmode='stack', hover_name=hover_name, hover_data={'Agenda Item': True},
                 custom_data=['Agenda Item'])
    fig.update_traces(meta={'filterDimension': 'ai'})
    fig.update_xaxes(type='category', categoryorder='total descending', title="Agenda Item / Topic")
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    if save_html:
        fig.write_html(str(export_dir / f"{prefix_id}_AI_Status.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_AI_Status'}}

    pivot = plot_df.pivot_table(index='Agenda Item', columns='TDoc Status', values='Count', aggfunc='sum', fill_value=0)
    pivot['Total'] = pivot.sum(axis=1)
    pivot = pivot.sort_values('Total', ascending=False).reset_index()
    for col in ['AI_Acronym', 'AI_Topic']:
        if col in valid_df.columns:
            meta = valid_df[['Agenda Item', col]].drop_duplicates('Agenda Item')
            pivot = pivot.merge(meta, on='Agenda Item', how='left')
    ordered = ['Agenda Item'] + [c for c in ['AI_Acronym', 'AI_Topic'] if c in pivot.columns] + [c for c in pivot.columns if c not in {'Agenda Item','AI_Acronym','AI_Topic'}]
    table = pivot[ordered].rename(columns={'AI_Acronym': 'Acronym', 'AI_Topic': 'Topic'})

    return (fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                        config=svg_config), dataframe_to_html_table(table))
