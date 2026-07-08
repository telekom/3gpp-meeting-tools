from plotly import express


def _generate_timeline(THEME_COLOR, svg_config, df, prefix, include_plotlyjs):
    valid_df = df.dropna(subset=['date_received']).sort_values('date_received').copy()
    if valid_df.empty: return "<p style='padding:20px; color:#666;'>No valid temporal data available.</p>"

    fig = px.histogram(
        valid_df, x='date_received', title="Email Traffic Over Time (1-Hour Bins)",
        color_discrete_sequence=[THEME_COLOR]
    )

    fig.update_traces(xbins=dict(size=3600000))
    fig.update_xaxes(rangeslider_visible=True, title="Timeline", tickformat="%a, %b %d<br>%H:%M",
                     ticklabelmode="period")
    fig.update_yaxes(title="Email Volume")

    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)