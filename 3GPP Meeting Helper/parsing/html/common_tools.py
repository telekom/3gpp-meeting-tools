import re
from typing import NamedTuple


def parse_tdoc_comments(comments, ignore_from_previous_meetings=True):
    if (comments is None) or (comments == ''):
        return TdocComments('', '', '', '')

    merge_of = ''
    merged_to = ''
    revision_of = ''
    revised_to = ''

    # Initial cleanup
    comments = comments.replace(', ', '').replace(',', '')

    # Comment parsing
    merge_of_match = merge_of_regex.search(comments)
    if merge_of_match is not None:
        str_full = merge_of_match[0]
        str_tdocs = merge_of_match.groupdict()['tdocs']
        comments = replace_and_clean_string(comments, str_full)

        str_tdocs_match = tdoc_regex.findall(str_tdocs)
        if str_tdocs_match is not None:
            merge_of = ', '.join(str_tdocs_match)

    revised_to_match = revised_to_regex.search(comments)
    if revised_to_match is not None:
        str_full = revised_to_match[0]
        comments = replace_and_clean_string(comments, str_full)
        revised_to = revised_to_match.groupdict()['tdoc']

    revision_of_match = revision_of_regex.search(comments)
    if revision_of_match is not None:
        str_full = revision_of_match[0]
        comments = replace_and_clean_string(comments, str_full)
        # Actually a XOR
        a = revision_of_match.groupdict()['previous_meeting'] is None
        b = ignore_from_previous_meetings
        if (a and b) or (not a and not b):
            revision_of = revision_of_match.groupdict()['tdoc']

    merged_to_match = merged_to_regex.search(comments)
    if merged_to_match is not None:
        str_full = merged_to_match[0]
        str_tdocs = merged_to_match.groupdict()['tdocs']
        replace_and_clean_string(comments, str_full)

        str_tdocs_match = tdoc_regex.findall(str_tdocs)
        if str_tdocs_match is not None:
            merged_to = ', '.join(str_tdocs_match)

    return TdocComments(revision_of, revised_to, merge_of, merged_to)


merged_to_regex = re.compile(r'[mM]erged (into|with) (?P<tdocs>(( and )?(, )?(part[s]? of)?( )?[S\d]*-\d\d[\d]+)+)')
revision_of_regex = re.compile(r'Revision of( Postponed)? (?P<tdoc>[S\d-]*)( from (?P<previous_meeting>S[AWG ]*[\d\w#-]*))?')
revised_to_regex = re.compile(
    r'Revised([ ]?(off-line|in parallel session|in drafting session))?(merging related CRs)?[ ]?to (?P<tdoc>[S\d-]*)')
merge_of_regex = re.compile(r'merging (?P<tdocs>(( and )?(, )?(part[s]? of)?( )?[S\d]*-\d\d[\d]+)+)')


def replace_and_clean_string(full_str, str_to_remove):
    return full_str.replace(str_to_remove, ' ').replace('  ', ' ')


tdoc_regex_str = r'S[\d]-\d\d\d[\d]+'
tdoc_regex = re.compile(tdoc_regex_str)


class TdocComments(NamedTuple):
    """
    Represents the comments of a parsed TDoc entry (revision of, revised to, etc.)
    """
    revision_of: str
    revised_to: str
    merge_of: str
    merged_to: str
