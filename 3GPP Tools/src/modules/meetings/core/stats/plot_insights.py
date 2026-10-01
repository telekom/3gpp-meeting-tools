import re
import pandas as pd
import plotly.express as px

from .stats_html import dataframe_to_html_table


def _is_revision(row) -> bool:
    parent = str(row.get('Is revision of', '') or '').strip()
    if parent and parent.lower() not in {'none', 'nan', '-'}:
        return True
    tdoc = str(row.get('TDoc', '') or '').strip()
    return bool(re.search(r'(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc, re.IGNORECASE))


def build_revision_stats(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    work['_IsRevision'] = work.apply(_is_revision, axis=1)
    rows = []
    for ai, group in work.groupby(work['Agenda Item'].fillna('').astype(str).str.strip()):
        if not ai:
            continue
        revisions = int(group['_IsRevision'].sum())
        originals = int((~group['_IsRevision']).sum())
        total = originals + revisions
        intensity = round((revisions / originals) * 100.0, 1) if originals else (100.0 if revisions else 0.0)
        first = group.iloc[0]
        rows.append({
            'AI': ai,
            'Acronym': str(first.get('AI_Acronym', '') or ''),
            'Topic': str(first.get('AI_Topic', '') or ''),
            'Original TDocs': originals,
            'Revisions': revisions,
            'Total': total,
            'Revisions / Originals %': intensity,
            '_Axis': str(first.get('AI_Axis', ai) or ai),
        })
    result = pd.DataFrame(rows)
    if not result.empty:
        result = result.sort_values(['Revisions', 'Total'], ascending=[False, False]).reset_index(drop=True)
    return result


def generate_revision_intensity_plot(df, export_dir, theme_color, prefix_id='Global', save_html=False):
    stats = build_revision_stats(df)
    if stats.empty or int(stats['Revisions'].sum()) == 0:
        msg = "<p style='padding:20px; color:#666;'>No revision lineage is available for this scope.</p>"
        return msg, dataframe_to_html_table(stats.drop(columns=['_Axis'], errors='ignore'))

    plot_df = stats.head(25).sort_values(['Revisions', 'Total'], ascending=True)
    fig = px.bar(
        plot_df,
        x=['Original TDocs', 'Revisions'],
        y='_Axis',
        orientation='h',
        barmode='stack',
        title='Revision Activity by Agenda Item',
        labels={'value': 'TDocs', '_Axis': 'Agenda Item / Topic', 'variable': 'Document Type'},
    )
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend_title_text='')
    fig.update_yaxes(title=None)
    if save_html:
        fig.write_html(str(export_dir / f'{prefix_id}_Revision_Activity.html'))
    table_df = stats.drop(columns=['_Axis'], errors='ignore')
    return fig.to_html(full_html=False, include_plotlyjs=False, default_height='100%', default_width='100%'), dataframe_to_html_table(table_df)


def build_company_specialization(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    company_docs = {}
    company_topics = {}
    for _, row in df.iterrows():
        companies = row.get('Clean_Companies', []) or []
        ai = str(row.get('Agenda Item', '') or '').strip()
        topic = str(row.get('AI_Topic', '') or '').strip() or str(row.get('AI_Acronym', '') or '').strip() or ai or 'Unspecified'
        tdoc = str(row.get('TDoc', '') or '').strip()
        for company in set(companies):
            company_docs.setdefault(company, set()).add(tdoc or f'row_{_}')
            company_topics.setdefault(company, {}).setdefault(topic, set()).add(tdoc or f'row_{_}')
    for company, docs in company_docs.items():
        topic_map = company_topics.get(company, {})
        if not topic_map:
            continue
        top_topic, top_docs = max(topic_map.items(), key=lambda kv: len(kv[1]))
        total = len(docs)
        top_count = len(top_docs)
        rows.append({
            'Company': company,
            'Total TDocs': total,
            'Topics': len(topic_map),
            'Top Topic': top_topic,
            'Top Topic TDocs': top_count,
            'Top Topic Share %': round((top_count / total) * 100.0, 1) if total else 0.0,
        })
    result = pd.DataFrame(rows)
    if not result.empty:
        result = result.sort_values(['Total TDocs', 'Top Topic Share %'], ascending=[False, False]).reset_index(drop=True)
    return result


def generate_company_specialization_plot(df, export_dir, theme_color, prefix_id='Global', save_html=False):
    stats = build_company_specialization(df)
    if stats.empty:
        msg = "<p style='padding:20px; color:#666;'>No company specialization data available.</p>"
        return msg, dataframe_to_html_table(stats)
    plot_df = stats.head(30).sort_values('Top Topic Share %', ascending=True)
    fig = px.scatter(
        plot_df,
        x='Top Topic Share %',
        y='Company',
        size='Total TDocs',
        hover_data=['Total TDocs', 'Topics', 'Top Topic', 'Top Topic TDocs'],
        custom_data=['Company'],
        title='Company Topic Specialization',
        labels={'Top Topic Share %': 'Share of TDocs in Top Topic (%)'},
        size_max=34,
    )
    fig.update_traces(meta={'filterDimension': 'company'})
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    fig.update_yaxes(title=None)
    if save_html:
        fig.write_html(str(export_dir / f'{prefix_id}_Company_Specialization.html'))
    return fig.to_html(full_html=False, include_plotlyjs=False, default_height='100%', default_width='100%'), dataframe_to_html_table(stats)
