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

    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Top_Contributors'}}

    return fig_comp.to_html(full_html=False, include_plotlyjs=False,
                            default_height="100%", default_width="100%", config=svg_config), len(comp_counts)


def generate_company_ai_heatmap(df, export_dir, prefix_id="Global", save_html=False):
    # Explode companies so co-signers get properly mapped to the matrix
    exploded_df = df.explode('Clean_Companies')
    exploded_df = exploded_df.dropna(subset=['Clean_Companies', 'Agenda Item'])
    exploded_df = exploded_df[(exploded_df['Clean_Companies'].str.strip() != '') &
                              (exploded_df['Agenda Item'].str.strip() != '')]

    # Target Top 20 Companies and Top 20 AIs to prevent an illegible massive matrix
    top_comps = exploded_df['Clean_Companies'].value_counts().head(25).index
    top_ais = exploded_df['Agenda Item'].value_counts().head(25).index

    plot_df = exploded_df[exploded_df['Clean_Companies'].isin(top_comps) &
                          exploded_df['Agenda Item'].isin(top_ais)]

    # Pivot to create the frequency matrix
    matrix = pd.crosstab(plot_df['Clean_Companies'], plot_df['Agenda Item'])

    # Order rows/cols by total volume so the heaviest hitters align nicely
    matrix = matrix.loc[matrix.sum(axis=1).sort_values(ascending=False).index]
    matrix = matrix[matrix.sum(axis=0).sort_values(ascending=False).index]

    fig = px.imshow(matrix,
                    labels=dict(x="Agenda Item", y="Company", color="TDocs"),
                    x=matrix.columns,
                    y=matrix.index,
                    text_auto=True,
                    aspect="auto",
                    title="Company Focus Matrix (Top 25 Companies vs Top 25 Topics)",
                    color_continuous_scale="Blues")

    fig.update_xaxes(side="bottom")
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    if save_html:
        fig.write_html(str(export_dir / f"{prefix_id}_Company_AI_Heatmap.html"))

    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Company_AI_Heatmap'}}

    return fig.to_html(full_html=False, include_plotlyjs=False,
                       default_height="100%", default_width="100%", config=svg_config)