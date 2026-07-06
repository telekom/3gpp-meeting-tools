# --- File: src/modules/meetings/core/stats/plot_contributors.py ---
import pandas as pd
import plotly.express as px


def generate_top_contributors_plot(df, export_dir, theme_color, top_count):
    # Extract and count all companies
    all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
    comp_counts = pd.Series(all_companies).value_counts().reset_index()
    comp_counts.columns = ['Company', 'Count']

    # Sort the top contributors so the highest is at the top of the chart
    plot_df = comp_counts.head(top_count).sort_values('Count', ascending=True)

    # Generate the bar chart
    fig_comp = px.bar(plot_df,
                      x='Count', y='Company', orientation='h',
                      title=f"Top {top_count} Contributing Companies",
                      color_discrete_sequence=[theme_color])

    # ---> THE FIX: Force Plotly to render every single company name on the Y-axis
    fig_comp.update_yaxes(tickmode='linear', dtick=1, title=None)

    # Write to standalone HTML and generate the div block for the dashboard
    fig_comp.write_html(str(export_dir / "Stat_Top_Contributors.html"))
    html_comp = fig_comp.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")

    return html_comp, len(comp_counts)