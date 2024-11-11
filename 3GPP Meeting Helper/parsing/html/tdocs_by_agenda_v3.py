import re
from typing import NamedTuple, List

import pandas as pd

from parsing.html.common_tools import parse_tdoc_comments


class HtmlSubstitution(NamedTuple):
    """
    Substitutions for th eHTML file
    """
    pattern: str
    repl: str
    flags: int
    name: str
    repetitions: int


def assert_if_tdocs_by_agenda_post_sa2_159(raw_html: str) -> bool:
    """
    Checks if the TDocs by Agenda is using the new format introduced with SA2#159 (October 2023)
    Args:
        raw_html: The raw HTML from TdocsByAgenda, e.g. https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_158_Goteborg_2023-08/TdocsByAgenda.htm

    Returns:
        bool: Whether the file uses the new format
    """
    if '<meta name="Generator" content="Microsoft Word 15 (filtered)">' in raw_html or '<meta name=Generator content="Microsoft Word 15 (filtered)">' in raw_html:
        return True
    return False


def apply_substitutions(raw_html: str, substitutions: List[HtmlSubstitution], logging=True)->str:
    original_size = len(raw_html)
    for idx, substitution in enumerate(substitutions):
        if substitution.repetitions == 0:
            raw_html = re.sub(pattern=substitution.pattern, repl=substitution.repl, string=raw_html,
                              flags=substitution.flags)
        else:
            raw_html = re.sub(pattern=substitution.pattern, repl=substitution.repl, string=raw_html,
                              flags=substitution.flags, count=substitution.repetitions)
        new_size = len(raw_html)
        if logging:
            print(f"Cleanup round {idx} ({substitution[3]}): HTML size={new_size} ({new_size / original_size * 100:.2f}%)")

    return raw_html


def parse_tdocs_by_agenda_v3(raw_html: str) -> pd.DataFrame:
    """
    Parses a TDocsByAgenda file from SA2#159 onwards
    Args:
        raw_html: The input document
    """

    # To do: enhance with https://webscraping.ai/faq/beautiful-soup/can-beautiful-soup-work-with-malformed-or-broken-html-xml-documents

    # some cleanup steps
    original_size = len(raw_html)
    print(f"Original HTML size={original_size}")
    substitutions = [
        HtmlSubstitution(r"&nbsp;", "", 0, 'White spaces', 0),
        HtmlSubstitution(r'<font( [\w]+=([#\w\d\-\:\"]*)?)*>', "", re.IGNORECASE, 'Comments', 0),
        HtmlSubstitution(r'<td( [\w]+( )?=[#\d\w]*)*>', "<TD>", re.IGNORECASE, 'Comments', 0),
        HtmlSubstitution(r'<b> Comment </b></FONT></TD>', "<b> Comments </b></FONT></TD>", re.IGNORECASE, 'Comments header', 0),
        HtmlSubstitution(r"<!--.*-->", "", re.DOTALL, 'Comments', 0),
        HtmlSubstitution(r"style='[^>']*'", "", re.IGNORECASE, 'style tags', 0),
        HtmlSubstitution(r"class=[^>]*", "", re.IGNORECASE, 'class tags"', 0),
        HtmlSubstitution(r"width=[^>]*", "", re.IGNORECASE, 'width tags"', 0),
        HtmlSubstitution(r"valign=[^>]*", "", re.IGNORECASE, 'valign tags"', 0),
        HtmlSubstitution(r"lang=[^>]*", "", re.IGNORECASE, 'lang tags"', 0),
        HtmlSubstitution(r"border='[^>']*'", "", re.IGNORECASE, 'border tags"', 0),
        HtmlSubstitution(r"cellpadding|cellspacing=[\d]+", "", re.IGNORECASE, 'cellpadding tags"', 0),
        HtmlSubstitution(r'<[/]?[pb][ ]*>', "", re.IGNORECASE, 'p, b tags', 0),
        HtmlSubstitution(r'<[/]?span[ ]*>', "", re.IGNORECASE, 'span tags', 0),
        HtmlSubstitution(r"&#39;", "'", 0, 'Apostrophe', 0),
        HtmlSubstitution(r"&amp;", "&", 0, 'Ampersand', 0),
    ]
    raw_html = apply_substitutions(raw_html, substitutions)
    tdoc_table = re.split(pattern="<table[ ]*>", string=raw_html, flags=re.IGNORECASE)[-1]

    second_substitutions = [
        HtmlSubstitution(r'[ ]{2,}', "", 0, 'Multiple spaces', 0),
        HtmlSubstitution(r' >', ">", 0, 'Spaces after >', 0),
        HtmlSubstitution(r'<[/]?thead>', "", re.IGNORECASE, 'Table header', 0),
        HtmlSubstitution(r'<[/]?tbody>', "", re.IGNORECASE, 'Table body', 0),
        HtmlSubstitution(r'<a href="[^"]*">', "", re.IGNORECASE, 'TDoc links opening tags', 0),
        HtmlSubstitution(r'</a>', "", re.IGNORECASE, 'TDoc links closing tags', 0),
        HtmlSubstitution(r'</font>', "", re.IGNORECASE, 'TDoc links closing tags', 0),
        HtmlSubstitution(r'<(th)( [\w]+( )?=[#\d\w]*)*>', "<td>", re.IGNORECASE, 'Header column separator (old)', 0)
    ]

    tdoc_table = apply_substitutions(tdoc_table, second_substitutions)

    rows_divided = [row for row in re.split(pattern=f"<tr( [\w]+=[\w]+)?>", string=tdoc_table, flags=re.IGNORECASE)]
    rows_divided = [row for row in rows_divided if row is not None]

    row_substitutions = [
        HtmlSubstitution(r'</(td|table|tr|th)>', "", re.IGNORECASE, 'End tags', 0),
        HtmlSubstitution(r'<br>', " ", re.IGNORECASE, 'End tags', 0),
        HtmlSubstitution(r'<span[\r\n]*>', "", re.IGNORECASE, 'Span', 0),
        HtmlSubstitution(r'\r\n', "", re.IGNORECASE, 'New lines', 0),
        HtmlSubstitution(r'<a.*href( )?=\"[^\"]*\">', "", re.IGNORECASE, 'Links', 0),
    ]
    rows = [[apply_substitutions(e, row_substitutions, logging=False).strip() for e in
             re.split(pattern=r"<td>", flags=re.IGNORECASE, string=row)
             if e.strip() != '']
            for row in rows_divided]

    def check_row(row):
        if row is None or row == [] or not isinstance(row, list) or len(row) < 2:
            return False

        # print(row)
        if row[1] == '-':
            return False
        return True
    rows = [row for row in
            rows
            if check_row(row) ]

    # print(f'TDocsByAgenda rows:\n{rows}')

    # for row in rows:
    #     print(row)

    title_row = rows[0]
    tdoc_rows = rows[1:]

    df_tdocs = pd.DataFrame(data=tdoc_rows, columns=title_row)
    df_tdocs = df_tdocs.set_index('TD#')
    if 'Subject' in df_tdocs:
        df_tdocs = df_tdocs.rename(columns={'Subject': 'Title'})

    df_tdocs['Revision of'] = ''
    df_tdocs['Revised to'] = ''
    df_tdocs['Merge of'] = ''
    df_tdocs['Merged to'] = ''
    df_tdocs['e-mail_Discussion'] = None

    tdoc_comments = {row[1]: parse_tdoc_comments(row[-2]) for row in tdoc_rows if row[-2] != ''}

    return df_tdocs
