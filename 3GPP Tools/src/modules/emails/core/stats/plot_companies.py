import pandas as pd
from plotly import express as px


def _generate_company_volume(THEME_COLOR, svg_config, df, prefix, include_plotlyjs, top_count):
    all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
    if not all_companies: return ""

    comp_counts = pd.Series(all_companies).value_counts().reset_index()
    comp_counts.columns = ['Company', 'Emails']

    # Use the dynamic top_count from config!
    plot_df = comp_counts.head(top_count).sort_values('Emails', ascending=True)

    fig = px.bar(plot_df, x='Emails', y='Company', orientation='h', title=f"Top {top_count} Active Companies",
                 color_discrete_sequence=[THEME_COLOR])

    fig.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1, title=None)

    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)


def _generate_company_ai_heatmap(svg_config, df, prefix, include_plotlyjs, top_comps_count, top_ais_count):
    # Explode lists so we can map individual companies to individual AIs
    exploded_df = df.explode('Clean_Companies').explode('ai_list')
    valid_df = exploded_df.dropna(subset=['Clean_Companies', 'ai_list'])
    valid_df = valid_df[(valid_df['Clean_Companies'].str.strip() != '') & (valid_df['ai_list'].str.strip() != '')]
    if valid_df.empty: return ""

    # Target Top N Companies and Top N AIs via dynamic variables
    top_comps = valid_df['Clean_Companies'].value_counts().head(top_comps_count).index
    top_ais = valid_df['ai_list'].value_counts().head(top_ais_count).index

    plot_df = valid_df[valid_df['Clean_Companies'].isin(top_comps) & valid_df['ai_list'].isin(top_ais)]
    if plot_df.empty: return ""

    # Pivot to create the frequency matrix
    matrix = pd.crosstab(plot_df['Clean_Companies'], plot_df['ai_list'])
    matrix = matrix.loc[matrix.sum(axis=1).sort_values(ascending=False).index]
    matrix = matrix[matrix.sum(axis=0).sort_values(ascending=False).index]

    fig = px.imshow(matrix, labels=dict(x="Agenda Item", y="Company", color="Emails"),
                    x=matrix.columns, y=matrix.index, text_auto=True, aspect="auto",
                    title=f"Company Focus Matrix (Top {top_comps_count} Companies vs Top {top_ais_count} Topics)",
                    color_continuous_scale="Blues")

    # ---> THE FIX: Force readable font sizes for the axes and the numbers inside the cells
    fig.update_yaxes(tickmode='linear', dtick=1, tickfont=dict(size=12))
    fig.update_xaxes(side="bottom", tickmode='linear', dtick=1, tickfont=dict(size=12))
    fig.update_traces(textfont=dict(size=13, weight='bold'))  # Forces the cell numbers to be larger and bold

    # Give the chart a bit more margin space so the larger labels don't get clipped off the edges
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=150, r=20, t=60, b=100)
    )

    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)