# --- File: src/modules/meetings/core/stats/plot_contributors.py ---
import pandas as pd
import plotly.express as px


def generate_top_contributors_plot(df, export_dir, theme_color, top_count, prefix_id="Global", save_html=False):
    all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
    comp_counts = pd.Series(all_companies).value_counts().reset_index()
    comp_counts.columns = ['Company', 'Count']

    plot_df = comp_counts.head(top_count).sort_values('Count', ascending=True)

    fig_comp = px.bar(plot_df, x='Count', y='Company', orientation='h',
                      title=f"Top {top_count} Contributing Companies",
                      color_discrete_sequence=[theme_color])

    fig_comp.update_yaxes(tickmode='linear', dtick=1, title=None)
    fig_comp.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1, title=None)

    fig_comp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    if save_html:
        fig_comp.write_html(str(export_dir / f"{prefix_id}_Top_Contributors.html"))

    html_comp = fig_comp.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    # ---> THE FIX: Force SVG and dynamic filename
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Top_Contributors'}}

    html_comp = fig_comp.to_html(full_html=False, include_plotlyjs=False,
                                 default_height="100%", default_width="100%", config=svg_config)

    return html_comp, len(comp_counts)