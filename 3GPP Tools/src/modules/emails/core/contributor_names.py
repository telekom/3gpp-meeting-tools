import re

import pandas as pd

from core.config.source_companies import SIGNATURE_SYNONYMS_REGEX, get_matching_contributors

recognized_vendor_list = [key for key, value in SIGNATURE_SYNONYMS_REGEX.items()]
recognized_vendor_list.append('Others')

contributor_columns = {}
for contributor in recognized_vendor_list:
    contributor_columns[contributor] = ('Contributed by ' + contributor)

others_cosigners: set[str] = set()


def reset_others():
    global others_cosigners
    others_cosigners = set()


ls_regex = re.compile(r'(.*) \((.*)\)')

# Removes unwanted characters and mentions to TDocs (e.g. in LS TDocs)
source_replace_regex = re.compile(r'\(Rapporteur\)|\(\?\)|[\[\]\?\(\)]|(([\w\d]{2,3})-(\d\d)([\d]+))')


def get_contributor_columns():
    return [value for key, value in contributor_columns.items()]


def add_contributor_columns_to_tdoc_list(df: pd.DataFrame, meeting_folder: str):
    """

    Args:
        df: The DataFrame where the list of contributor columns is to be added
        meeting_folder: The meeting folder (used only for logginc)

    Returns:
        The co-signers that could not be assigned and the DataFrame with the added columns
    """
    print('Adding contributor columns for meeting folder {0}'.format(meeting_folder))
    # start = time.perf_counter()

    others_cosigners = set()
    known_cosigners = set()

    df['Source (summary)'] = ''
    all_contributor_columns = get_contributor_columns()

    # More efficient alternative than df[all_contributor_columns] = False (threw a performance warning). Not a
    # super-big-deal to go from 145ms to 100ms, but the warnings were annoying...
    # df[all_contributor_columns] = False
    df_data_to_concatenate = [df]
    for new_col_name in all_contributor_columns:
        df_data_to_concatenate.append(pd.Series(False, name=new_col_name, index=df.index))
    df = pd.concat(df_data_to_concatenate, axis=1)

    for tdoc in df.index:
        tdoc_source = df.at[tdoc, 'Source']
        found_cosigners = get_matching_contributors(tdoc_source, others_cosigners, known_cosigners)

        # Fill in the summary contributor columns
        for cosigner in found_cosigners:
            contributor_column = contributor_columns[cosigner]
            df.at[tdoc, contributor_column] = True
        # Summary column
        if len(found_cosigners) > 0:
            df.at[tdoc, 'Source (summary)'] = ', '.join(found_cosigners)

    # end = time.perf_counter()
    # ms = (end - start) * 10 ** 3
    # print(f"Elapsed {ms:.03f} ms.")

    # others_cosigners contains all the cosigners that could not be mapped to a source
    return others_cosigners, df
