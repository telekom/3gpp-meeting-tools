def _generate_delegate_table(df, prefix):
    delegates = df.groupby('sender_email').agg(
        Name=('sender_name', lambda x: x.value_counts().index[0] if not x.empty else "Unknown"),
        Company=('Clean_Companies', lambda x: x.iloc[0][0] if len(x.iloc[0]) > 0 else "Unknown"),
        Emails=('id', 'count')
    ).reset_index().sort_values('Emails', ascending=False)

    rows = ""
    for _, row in delegates.iterrows():
        rows += f"<tr><td>{row['Name']}</td><td>{row['sender_email']}</td><td>{row['Company']}</td><td>{row['Emails']}</td></tr>"

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