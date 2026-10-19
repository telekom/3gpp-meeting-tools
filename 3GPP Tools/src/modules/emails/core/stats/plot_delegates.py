import html
from plotly import express as px


def _generate_delegates_plot(THEME_COLOR, svg_config, df, prefix, include_plotlyjs, top_count):
    # Explode by company and drop missing emails
    exploded_df = df.explode('Clean_Companies')
    valid_df = exploded_df.dropna(subset=['Clean_Companies', 'sender_email'])
    valid_df = valid_df[(valid_df['Clean_Companies'].str.strip() != '') & (valid_df['sender_email'].str.strip() != '')]
    if valid_df.empty: return ""

    # Count unique sender emails per company
    delegates_count = valid_df.groupby('Clean_Companies')['sender_email'].nunique().reset_index()
    delegates_count.columns = ['Company', 'Active Delegates']

    plot_df = delegates_count.sort_values('Active Delegates', ascending=False).head(top_count).sort_values(
        'Active Delegates', ascending=True)

    fig = px.bar(plot_df, x='Active Delegates', y='Company', orientation='h',
                 title=f"Top {top_count} Companies by Active Delegates", color_discrete_sequence=[THEME_COLOR])

    fig.update_yaxes(type='category', categoryorder='total ascending', tickmode='linear', dtick=1, title=None)
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    return fig.to_html(full_html=False, include_plotlyjs=include_plotlyjs,
                       default_height="100%", default_width="100%", config=svg_config)


def _first_company(series):
    for comps in series:
        if isinstance(comps, (list, tuple)):
            for c in comps:
                if str(c).strip():
                    return str(c).strip()
        elif isinstance(comps, str) and comps.strip():
            return comps.strip()
    return "Unknown"


def _generate_delegate_table(df, prefix):
    delegates = df.groupby('sender_email').agg(
        Name=('sender_name', lambda x: x.value_counts().index[0] if not x.empty else "Unknown"),
        Company=('Clean_Companies', _first_company),
        Emails=('id', 'count')
    ).reset_index().sort_values('Emails', ascending=False)

    rows = ""
    for _, row in delegates.iterrows():
        safe_name = html.escape(str(row['Name']))
        safe_email = html.escape(str(row['sender_email']))
        safe_company = html.escape(str(row['Company']))
        rows += f"<tr><td>{safe_name}</td><td>{safe_email}</td><td>{safe_company}</td><td>{int(row['Emails'])}</td></tr>"

    safe_id = f"table_{prefix}".replace(" ", "_").replace(".", "_")

    table_template = """
    <h3 style="color:#333; text-align:center;">Top Active Delegates</h3>
    <table id="__SAFE_ID__" class="display delegate-table" style="width:100%">
        <thead><tr><th>Name</th><th>Email Address</th><th>Company</th><th>Sent Emails</th></tr></thead>
        <tbody>__ROWS__</tbody>
    </table>
    """
    table_template = table_template.replace("__SAFE_ID__", str(safe_id))
    table_template = table_template.replace("__ROWS__", str(rows))

    return table_template