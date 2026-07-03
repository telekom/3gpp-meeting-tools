import plotly.express as px

def generate_ai_volume_plot(df, export_dir, theme_color):
    ai_counts = df['Agenda Item'].value_counts().reset_index()
    ai_counts.columns = ['Agenda Item', 'Count']
    ai_counts = ai_counts[ai_counts['Agenda Item'].str.strip() != '']
    fig_ai = px.bar(ai_counts.head(20), x='Agenda Item', y='Count', title="Top 20 Agenda Items by TDoc Volume",
                    color_discrete_sequence=[theme_color])
    fig_ai.write_html(str(export_dir / "Stat_AI_Volume.html"))
    return fig_ai.to_html(full_html=False, include_plotlyjs='cdn', default_height="100%", default_width="100%")