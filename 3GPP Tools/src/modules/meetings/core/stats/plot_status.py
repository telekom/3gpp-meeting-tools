import pandas as pd
import plotly.express as px
from .stats_html import dataframe_to_html_table


def generate_outcomes_plot(df, export_dir, palette, prefix_id="Global", save_html=False):
    status_counts = df['TDoc Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    status_counts = status_counts[status_counts['Status'].astype(str).str.strip() != '']
    total = int(status_counts['Count'].sum())
    status_counts['Percentage'] = status_counts['Count'].apply(lambda x: f"{(100*x/total):.1f}%" if total else "0.0%")

    fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.4,
                        title="TDoc Outcomes", color_discrete_sequence=palette, custom_data=['Status'])
    fig_status.update_traces(meta={'filterDimension': 'status'})
    fig_status.update_traces(textposition='inside', textinfo='percent+label')
    fig_status.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    if save_html:
        fig_status.write_html(str(export_dir / f"{prefix_id}_Outcomes.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Outcomes'}}
    table = status_counts.rename(columns={'Count': 'TDocs'})
    return (fig_status.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                               config=svg_config), dataframe_to_html_table(table))
