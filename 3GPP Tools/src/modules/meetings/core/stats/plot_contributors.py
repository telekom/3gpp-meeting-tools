import pandas as pd
import plotly.express as px
from .stats_html import dataframe_to_html_table


def generate_top_contributors_plot(df, export_dir, theme_color, top_count, prefix_id="Global", save_html=False):
    all_companies = [comp for sublist in df['Clean_Companies'] for comp in sublist]
    comp_counts = pd.Series(all_companies).value_counts().reset_index()
    comp_counts.columns = ['Company', 'Count']
    top_df = comp_counts.head(top_count)
    plot_df = top_df.sort_values('Count', ascending=True)
    fig_comp = px.bar(plot_df, x='Count', y='Company', orientation='h', title=f"Top {top_count} Contributing Companies",
                      color_discrete_sequence=[theme_color], custom_data=['Company'])
    fig_comp.update_traces(meta={'filterDimension': 'company'})
    fig_comp.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1, title=None)
    fig_comp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    if save_html:
        fig_comp.write_html(str(export_dir / f"{prefix_id}_Top_Contributors.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Top_Contributors'}}
    return (fig_comp.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                             config=svg_config), len(comp_counts),
            dataframe_to_html_table(top_df.rename(columns={'Count': 'TDocs'})))


def generate_company_ai_heatmap(df, export_dir, prefix_id="Global", save_html=False, top_comps_count=25, top_ais_count=25):
    axis_col = 'AI_Axis' if 'AI_Axis' in df.columns else 'Agenda Item'
    exploded_df = df.explode('Clean_Companies', ignore_index=True).dropna(subset=['Clean_Companies', 'Agenda Item']).copy()
    exploded_df['Clean_Companies'] = exploded_df['Clean_Companies'].astype(str).str.strip()
    exploded_df['Agenda Item'] = exploded_df['Agenda Item'].astype(str).str.strip()
    exploded_df[axis_col] = exploded_df[axis_col].astype(str).str.strip()
    exploded_df = exploded_df[(exploded_df['Clean_Companies'] != '') & (exploded_df['Agenda Item'] != '')].reset_index(drop=True)

    top_comps = exploded_df['Clean_Companies'].value_counts().head(top_comps_count).index
    top_ais = exploded_df['Agenda Item'].value_counts().head(top_ais_count).index
    plot_df = exploded_df[exploded_df['Clean_Companies'].isin(top_comps) & exploded_df['Agenda Item'].isin(top_ais)]
    matrix = pd.crosstab(plot_df['Clean_Companies'], plot_df[axis_col])
    matrix = matrix.loc[matrix.sum(axis=1).sort_values(ascending=False).index]
    matrix = matrix[matrix.sum(axis=0).sort_values(ascending=False).index]

    fig = px.imshow(matrix, labels=dict(x="Agenda Item / Topic", y="Company", color="TDocs"), x=matrix.columns,
                    y=matrix.index, text_auto=True, aspect="auto",
                    title=f"Company Focus Matrix (Top {top_comps_count} Companies vs Top {top_ais_count} Topics)",
                    color_continuous_scale="Blues")
    fig.update_yaxes(tickmode='linear', dtick=1, tickfont=dict(size=12))
    fig.update_xaxes(side="bottom", tickmode='linear', dtick=1, tickfont=dict(size=12))
    fig.update_traces(textfont=dict(size=13, weight='bold'))
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    if save_html:
        fig.write_html(str(export_dir / f"{prefix_id}_Company_AI_Heatmap.html"))
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Company_AI_Heatmap'}}
    table = matrix.reset_index().rename(columns={'Clean_Companies': 'Company'})
    return (fig.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%",
                        config=svg_config), dataframe_to_html_table(table))
