# --- File: src/modules/meetings/core/stats/plot_agenda.py ---
import plotly.express as px


def generate_ai_volume_plot(df, export_dir, theme_color, prefix_id="Global", save_html=False):
    ai_counts = df['Agenda Item'].value_counts().reset_index()
    ai_counts.columns = ['Agenda Item', 'Count']
    ai_counts = ai_counts[ai_counts['Agenda Item'].str.strip() != '']

    fig_ai = px.bar(ai_counts.head(20), x='Agenda Item', y='Count',
                    title="Top 20 Agenda Items by TDoc Volume",
                    color_discrete_sequence=[theme_color])

    fig_ai.update_xaxes(type='category', categoryorder='total descending')

    if save_html:
        fig_ai.write_html(str(export_dir / f"{prefix_id}_AI_Volume.html"))

    # ---> THE FIX: Force SVG and dynamic filename
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_AI_Volume'}}

    return fig_ai.to_html(full_html=False, include_plotlyjs='cdn',
                          default_height="100%", default_width="100%", config=svg_config)