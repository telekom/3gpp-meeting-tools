from plotly import express as px


def _generate_ai_volume(THEME_COLOR, svg_config, df, prefix, include_plotlyjs):
    exploded_df = df.explode('ai_list')
    valid_df = exploded_df.dropna(subset=['ai_list'])
    if valid_df.empty: return ""

    counts = valid_df['ai_list'].value_counts().reset_index()
    counts.columns = ['Agenda Item', 'Emails']

    fig = px.bar(counts, x='Agenda Item', y='Emails', title="Agenda Items by Email Volume",
                 color_discrete_sequence=[THEME_COLOR])

    fig.update_xaxes(type='category', categoryorder='total descending')

    # Matches plot_agenda.py flawlessly and assigns SVG as default export
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)