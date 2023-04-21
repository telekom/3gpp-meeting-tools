import os.path
from typing import Tuple, NamedTuple, List

import html2text
import pandas as pd
import re
import traceback


class TdocRevision(NamedTuple):
    """NamedTuple containing the TDoc ID and the revision number (plus a '*' character if it is a draft)"""
    tdoc: str
    revision: str


def extract_tdoc_revisions_from_html(
        html_content: str,
        is_draft=False,
        is_path=False,
        ignore_revision=False) -> List[TdocRevision]:
    """
    Extracts TDoc revisions from an HTML TDoc files being mentioned there
    Args:
        is_path: Whether html_content is actually a file path
        html_content: The HTML content to parse or a file path if is_path is True
        is_draft: Whether this file lists draft revisions (adds "*" to the revision name)
        ignore_revision: If True, only parses the TDoc number and ignores the revision (empty value)

    Returns: A tuple containing the TDoc number and the revision

    """
    if is_path:
        if not os.path.exists(html_content):
            return []
        try:
            with open(html_content, 'r') as file:
                html_content = file.read()
        except:
            print('Could not open file "{0}"'.format(html_content))
            return []

    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    tdoc_revisions_text = h.handle(html_content)
    if ignore_revision:
        tdocs = re.findall(r'S2-[\d]{7}', tdoc_revisions_text)
    else:
        tdocs = re.findall(r'S2-[\d]{7}r[\d]{2}', tdoc_revisions_text)
    tdocs = list(set(tdocs))

    print('Extracting TDoc revisions: text length={0}, {1} TDoc revisions found. ignore_revisions={2}, is_draft={3}'.format(
        len(tdoc_revisions_text),
        len(tdocs),
        ignore_revision,
        is_draft))

    if ignore_revision:
        tdoc_list = [TdocRevision(tdoc, '') for tdoc in tdocs]
    else:
        if not is_draft:
            tdoc_list = [TdocRevision(tdoc[0:-3], tdoc[-2:]) for tdoc in tdocs]
        else:
            tdoc_list = [TdocRevision(tdoc[0:-3], tdoc[-2:] + '*') for tdoc in tdocs]
    return tdoc_list

def revisions_file_to_dataframe(
        revisions_file: str,
        meeting_tdocs: pd.DataFrame,
        drafts_file=None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parses a revisions file HTML and extracts the list of TDoc revisions in the file
    Args:
        revisions_file: The location of the revisions file
        meeting_tdocs: A TDoc list to which to add in "Revisions" column the last revision
        drafts_file: An optional drafts file to extract

    Returns:

    """
    try:
        revision_list = extract_tdoc_revisions_from_html(revisions_file, is_path=True)

        try:
            if drafts_file is not None:
                print('Parsing drafts files: {0}'.format(drafts_file))
                drafts_list = extract_tdoc_revisions_from_html(drafts_file, is_draft=True, is_path=True)
                revision_list.extend(drafts_list)
        except:
            print('Could not open drafts file {0}'.format(drafts_file))
            # traceback.print_exc()

        df = pd.DataFrame(revision_list, columns=['Tdoc', 'Revisions'])
        # df["Revisions"] = df[["Revisions"]].apply(pd.to_numeric) # We now also have drafts
        # print(revision_list)

        df_per_tdoc = df.groupby("Tdoc")
        maximums = df_per_tdoc.max()
        maximums.sort_values(by='Revisions', ascending=False, inplace=True)
        maximums = maximums.reset_index()
        maximums = maximums.set_index('Tdoc')
        # display(maximums)

        meeting_tdocs = pd.concat([meeting_tdocs, maximums], axis=1, sort=False)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions'].fillna(0)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions']  # .astype(int)

        indexed_df = df.set_index('Tdoc')

        return meeting_tdocs, indexed_df
    except:
        print('Could not parse revisions file {0}. Drafts file: {1}'.format(revisions_file, drafts_file))
        # traceback.print_exc()
        return None, None
