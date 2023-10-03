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


def apply_substitutions(raw_html:str, substitutions: List[HtmlSubstitution]):
    original_size = len(raw_html)
    for idx, substitution in enumerate(substitutions):
        if substitution.repetitions == 0:
            raw_html = re.sub(pattern=substitution.pattern, repl=substitution.repl, string=raw_html,
                              flags=substitution.flags)
        else:
            raw_html = re.sub(pattern=substitution.pattern, repl=substitution.repl, string=raw_html,
                              flags=substitution.flags, count=substitution.repetitions)
        new_size = len(raw_html)
        print(f"Cleanup round {idx} ({substitution[3]}): HTML size={new_size} ({new_size / original_size * 100:.2f}%)")

    return raw_html


def parse_tdocs_by_agenda_v3(raw_html: str):
    """
    Parses a TDocsByAgenda file from SA2#159 onwards
    Args:
        raw_html: The input document
    """
    # some cleanup steps
    original_size = len(raw_html)
    print(f"Original HTML size={original_size}")
    substitutions = [
        HtmlSubstitution(r"&nbsp;", "", 0, 'White spaces', 0),
        HtmlSubstitution(r"<!--.*-->", "", re.DOTALL, 'Comments', 0),
        HtmlSubstitution(r"style='[^>']*'", "", 0, 'style tags', 0),
        HtmlSubstitution(r"class=[^>]*", "", 0, 'class tags"', 0),
        HtmlSubstitution(r"width=[^>]*", "", 0, 'width tags"', 0),
        HtmlSubstitution(r"valign=[^>]*", "", 0, 'valign tags"', 0),
        HtmlSubstitution(r"lang=[^>]*", "", 0, 'lang tags"', 0),
        HtmlSubstitution(r"border='[^>']*'", "", 0, 'border tags"', 0),
        HtmlSubstitution(r"cellpadding|cellspacing=[\d]+", "", 0, 'cellpadding tags"', 0),
        HtmlSubstitution(r'<[/]?[pb][ ]*>', "", 0, 'p, b tags', 0),
        HtmlSubstitution(r'<[/]?span[ ]*>', "", 0, 'span tags', 0),
    ]
    raw_html = apply_substitutions(raw_html, substitutions)
    tdoc_table = re.split(pattern="<table[ ]*>", string=raw_html)[-1]

    second_substitutions = [
        #HtmlSubstitution(r'[\n]', "", 0, 'Carriage Returns', 0),
        HtmlSubstitution(r'[ ]{2,}', "", 0, 'Multiple spaces', 0),
        HtmlSubstitution(r' >', ">", 0, 'Spaces after >', 0),
        HtmlSubstitution(r'<[/]?thead>', "", 0, 'Table header', 0),
        HtmlSubstitution(r'<[/]?tbody>', "", 0, 'Table body', 0),
        HtmlSubstitution(r'<a href="[^"]*">', "", 0, 'TDoc links opening tags', 0),
        HtmlSubstitution(r'</a>', "", 0, 'TDoc links closing tags', 0),
    ]
    tdoc_table = apply_substitutions(tdoc_table, second_substitutions)

    rows = [[e.replace('</td>', '').replace('</tr>', '').replace('<span\r\n>', '').replace('\r\n', ' ').strip() for e in row.split("<td>") if e.strip() != ''] for row in tdoc_table.split("<tr>")]
    rows = [row for row in rows if row != [] and row[1]!='-']

    # for row in rows:
    #    print(row)

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