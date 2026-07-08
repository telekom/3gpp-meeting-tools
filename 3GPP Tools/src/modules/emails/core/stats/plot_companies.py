import pandas
from plotly import express


def _generate_company_volume(THEME_COLOR, svg_config, df, prefix, include_plotlyjs):
    all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
    if not all_companies: return ""

    comp_counts = pd.Series(all_companies).value_counts().reset_index()
    comp_counts.columns = ['Company', 'Emails']

    plot_df = comp_counts.head(25).sort_values('Emails', ascending=True)

    fig = px.bar(plot_df, x='Emails', y='Company', orientation='h', title="Top 25 Active Companies",
                 color_discrete_sequence=[THEME_COLOR])

    fig.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1, title=None)

    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)