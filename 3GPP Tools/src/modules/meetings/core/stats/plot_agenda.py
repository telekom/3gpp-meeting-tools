# --- File: src/modules/meetings/core/stats/plot_agenda.py ---
import plotly.express as px


def generate_ai_volume_plot(df, export_dir, theme_color, prefix_id="Global", save_html=False):
    ai_counts = df['Agenda Item'].value_counts().reset_index()
    ai_counts.columns = ['Agenda Item', 'Count']
    ai_counts = ai_counts[ai_counts['Agenda Item'].str.strip() != '']

    fig_ai = px.bar(ai_counts, x='Agenda Item', y='Count',
                    title="Agenda Items by TDoc Volume",
                    color_discrete_sequence=[theme_color])

    fig_ai.update_xaxes(type='category', categoryorder='total descending')
    fig_ai.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    if save_html:
        fig_ai.write_html(str(export_dir / f"{prefix_id}_AI_Volume.html"))

    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_AI_Volume'}}

    return fig_ai.to_html(full_html=False, include_plotlyjs='cdn',
                          default_height="100%", default_width="100%", config=svg_config)


def generate_ai_status_plot(df, export_dir, palette, prefix_id="Global", save_html=False):
    # Filter out missing AIs and statuses
    valid_df = df[(df['Agenda Item'].str.strip() != '') & (df['TDoc Status'].str.strip() != '')].copy()

    # Group by Agenda Item and Status
    counts = valid_df.groupby(['Agenda Item', 'TDoc Status']).size().reset_index(name='Count')

    # Get Top 20 AIs by total volume to keep the chart readable
    top_ais = valid_df['Agenda Item'].value_counts().index
    plot_df = counts[counts['Agenda Item'].isin(top_ais)]

    # barmode='stack' builds the layered visualization automatically
    fig = px.bar(plot_df, x='Agenda Item', y='Count', color='TDoc Status',
                 title="Agenda Items by Outcome Status",
                 color_discrete_sequence=palette,
                 barmode='stack')

    fig.update_xaxes(type='category', categoryorder='total descending')
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    if save_html:
        fig.write_html(str(export_dir / f"{prefix_id}_AI_Status.html"))

    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_AI_Status'}}

    return fig.to_html(full_html=False, include_plotlyjs=False,
                       default_height="100%", default_width="100%", config=svg_config)