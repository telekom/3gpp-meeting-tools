import plotly.express as px

def generate_outcomes_plot(df, export_dir, palette):
    status_counts = df['TDoc Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    status_counts = status_counts[status_counts['Status'].str.strip() != '']
    fig_status = px.pie(status_counts, names='Status', values='Count', hole=0.4, title="TDoc Outcomes",
                        color_discrete_sequence=palette)
    fig_status.update_traces(textposition='inside', textinfo='percent+label')
    fig_status.write_html(str(export_dir / "Stat_Outcomes.html"))
    return fig_status.to_html(full_html=False, include_plotlyjs=False, default_height="100%", default_width="100%")