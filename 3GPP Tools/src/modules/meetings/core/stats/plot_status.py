# --- File: src/modules/meetings/core/stats/plot_status.py ---
import plotly.express as px


def generate_outcomes_plot(df, export_dir, palette, prefix_id="Global", save_html=False):
    status_counts = df['TDoc Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    status_counts = status_counts[status_counts['Status'].str.strip() != '']

    fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.4,
                        title="TDoc Outcomes", color_discrete_sequence=palette)
    fig_status.update_traces(textposition='inside', textinfo='percent+label')

    if save_html:
        fig_status.write_html(str(export_dir / f"{prefix_id}_Outcomes.html"))

    # ---> THE FIX: Force SVG and dynamic filename
    svg_config = {'toImageButtonOptions': {'format': 'svg', 'filename': f'{prefix_id}_Outcomes'}}

    return fig_status.to_html(full_html=False, include_plotlyjs=False,
                              default_height="100%", default_width="100%", config=svg_config)